import sys
import os

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QTextEdit,
    QPushButton,
    QMessageBox,
    QDialog,
    QTabWidget,
)
from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal


# -------------------- Popup dialog --------------------

class MetadataDialog(QDialog):
    def __init__(self, variables_text: str, outputs_text: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("HFSS Metadata")
        self.resize(800, 600)

        tabs = QTabWidget()

        var_view = QTextEdit()
        var_view.setReadOnly(True)
        var_view.setText(variables_text)
        tabs.addTab(var_view, "Variables")

        out_view = QTextEdit()
        out_view.setReadOnly(True)
        out_view.setText(outputs_text)
        tabs.addTab(out_view, "Outputs")

        layout = QVBoxLayout()
        layout.addWidget(tabs)
        self.setLayout(layout)


# -------------------- Worker (threaded metadata fetch) --------------------

class MetadataFetchWorker(QObject):
    finished = pyqtSignal(object)   # dict payload
    error = pyqtSignal(str)

    def __init__(self, port: str, version: str, parent=None):
        super().__init__(parent)
        self.port = port
        self.version = version

    def run(self):
        """Connect to AEDT and fetch ONLY legacy-style metadata:
        - Variables: local design variables (oDesign.GetVariables)
        - Outputs: report names (ReportSetup.GetAllReportNames)
        """
        try:
            from ansys.aedt.core import Desktop
            from ansys.aedt.core.generic.design_types import get_pyaedt_app
        except Exception as e:
            self.error.emit(f"PyAEDT import failed: {e}")
            return

        if not self.port:
            self.error.emit(
                "PYAEDT_SCRIPT_PORT not found in environment.\n"
                "This GUI must be launched from the AEDT PyAEDT Extension button."
            )
            return

        desktop = None
        try:
            desktop = Desktop(
                version=self.version,
                new_desktop=False,
                port=int(self.port),
            )
            app = get_pyaedt_app(desktop=desktop)

            payload = {
                "variables": [],     # local design variables only
                "outputs": [],       # report names only
                "errors": [],        # non-fatal issues
            }

            # ---- Variables: local design variables (legacy oDesign.GetVariables) ----
            try:
                payload["variables"] = list(app.variable_manager.design_variable_names or [])
            except Exception as e:
                payload["errors"].append(f"Failed to read design variables: {e}")

            # ---- Outputs: report names (legacy ReportSetup.GetAllReportNames) ----
            try:
                payload["outputs"] = list(getattr(app.post, "all_report_names", []) or [])
            except Exception as e:
                payload["errors"].append(f"Failed to read report names: {e}")

            self.finished.emit(payload)

        except Exception as e:
            self.error.emit(f"Failed to fetch metadata: {e}")

        finally:
            if desktop:
                try:
                    desktop.release_desktop(close_projects=False, close_on_exit=False)
                except Exception:
                    pass


# -------------------- Main window --------------------

class MainWindow(QWidget):
    def __init__(self, project_path=None, parent=None):
        super().__init__(parent)
        self.project_path = project_path
        self.desktop = None
        self.app = None  # PyAEDT app (HFSS, etc.)

        # thread handles (keep refs)
        self._meta_thread = None
        self._meta_worker = None

        self.init_ui()

    # ---------- GUI setup ----------

    def init_ui(self):
        title = "Optimizer GUI"
        if self.project_path:
            title += f" - {self.project_path}"
        self.setWindowTitle(title)

        self.resize(600, 400)
        self.setAttribute(Qt.WA_DeleteOnClose, True)

        layout = QVBoxLayout()

        self.start_btn = QPushButton("Start Simulation")
        self.start_btn.clicked.connect(self.run_simulation)
        layout.addWidget(self.start_btn)

        self.fetch_meta_btn = QPushButton("Fetch Metadata")
        self.fetch_meta_btn.clicked.connect(self.fetch_metadata)
        layout.addWidget(self.fetch_meta_btn)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.setLayout(layout)

        self.append_log("Optimizer GUI started.")
        if self.project_path:
            self.append_log(f"Received AEDT project path: {self.project_path}")

    def append_log(self, text: str):
        self.log.append(text)

    # ---------- AEDT / PyAEDT helpers (existing) ----------

    def connect_to_open_aedt(self):
        """Attach to the currently open AEDT session using PyAEDT."""
        try:
            from ansys.aedt.core import Desktop
            from ansys.aedt.core.generic.design_types import get_pyaedt_app
        except ImportError as e:
            self.append_log("ERROR: PyAEDT is not installed in this Python environment.")
            self.append_log(str(e))
            return False

        port = os.environ.get("PYAEDT_SCRIPT_PORT")
        version = os.environ.get("PYAEDT_SCRIPT_VERSION", "2025.2")

        if not port:
            self.append_log("ERROR: PYAEDT_SCRIPT_PORT not found in environment.")
            self.append_log("This GUI must be launched from the AEDT PyAEDT Extension button.")
            return False

        self.append_log(f"Connecting to AEDT on gRPC port {port} (version={version})...")

        try:
            self.desktop = Desktop(
                version=version,
                new_desktop=False,
                port=int(port),
            )

            self.app = get_pyaedt_app(desktop=self.desktop)

            self.append_log(
                f"Connected to project '{self.app.project_name}', "
                f"design '{self.app.design_name}'."
            )
            return True

        except Exception as e:
            self.append_log(f"ERROR: Failed to attach to AEDT: {e}")
            self.desktop = None
            self.app = None
            return False

    def run_hfss_simulation_and_get_s11_summary(self):
        if not self.app:
            self.append_log("ERROR: PyAEDT app not available.")
            return None

        try:
            self.append_log("Starting HFSS simulation via app.analyze()...")
            ok = self.app.analyze()
            self.append_log(f"HFSS analyze() returned: {ok}")
            if not ok:
                self.append_log("WARNING: analyze() returned False (simulation may have failed).")
        except Exception as e:
            self.append_log(f"ERROR during analyze(): {e}")
            return None

        try:
            self.append_log("Retrieving dB(S(1,1)) solution data...")

            setup_sweep = self.app.nominal_sweep
            data = self.app.post.get_solution_data(
                expressions="S(1,1)",
                setup_sweep_name=setup_sweep,
            )

            freqs, s11_db = data.get_expression_data(
                expression="S(1,1)",
                formula="db20",
            )

            n = len(freqs)
            if n == 0:
                self.append_log("No data points returned from get_expression_data().")
                return None

            lines = [f"Retrieved {n} points of dB(S(1,1)) (showing first few):"]
            max_to_show = min(5, n)
            for i in range(max_to_show):
                lines.append(f"{i}: Freq = {freqs[i]}, dB(S11) = {s11_db[i]}")
            return "\n".join(lines)

        except Exception as e:
            self.append_log(f"ERROR retrieving solution data: {e}")
            return None

    # ---------- Simulation entry (existing) ----------

    def run_simulation(self):
        self.start_btn.setEnabled(False)
        try:
            self.append_log("Attempting to connect to open AEDT session...")
            if not self.connect_to_open_aedt():
                self.append_log("Aborting simulation because connection failed.")
                QMessageBox.critical(self, "Simulation Failed", "Failed to connect to AEDT.\nCheck the log for details.")
                return

            summary = self.run_hfss_simulation_and_get_s11_summary()
            if summary:
                QMessageBox.information(self, "Simulation Result", summary)
            else:
                QMessageBox.warning(self, "Simulation Result", "Simulation finished, but no result summary is available.\nCheck the log for details.")

        finally:
            if self.desktop:
                try:
                    self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                    self.append_log("Released PyAEDT Desktop connection.")
                except Exception as e:
                    self.append_log(f"WARNING: Failed to release Desktop cleanly: {e}")
            self.start_btn.setEnabled(True)

    # ---------- Fetch Metadata (new) ----------

    def fetch_metadata(self):
        port = os.environ.get("PYAEDT_SCRIPT_PORT")
        version = os.environ.get("PYAEDT_SCRIPT_VERSION", "2025.2")

        self.fetch_meta_btn.setEnabled(False)
        self.append_log("Fetching metadata from AEDT...")

        self._meta_thread = QThread()
        self._meta_worker = MetadataFetchWorker(port=port, version=version)
        self._meta_worker.moveToThread(self._meta_thread)

        self._meta_thread.started.connect(self._meta_worker.run)
        self._meta_worker.finished.connect(self._on_metadata_ready)
        self._meta_worker.error.connect(self._on_metadata_error)

        # cleanup
        self._meta_worker.finished.connect(self._meta_thread.quit)
        self._meta_worker.error.connect(self._meta_thread.quit)
        self._meta_worker.finished.connect(self._meta_worker.deleteLater)
        self._meta_worker.error.connect(self._meta_worker.deleteLater)
        self._meta_thread.finished.connect(self._meta_thread.deleteLater)

        self._meta_thread.start()

    def _on_metadata_error(self, msg: str):
        self.fetch_meta_btn.setEnabled(True)
        self.append_log(f"Metadata fetch failed: {msg}")
        QMessageBox.critical(self, "Fetch Metadata Failed", msg)

    def _on_metadata_ready(self, payload: dict):
        self.fetch_meta_btn.setEnabled(True)
        self.append_log("Metadata fetch completed.")

        design_vars = payload.get("variables", []) or []
        report_names = payload.get("outputs", []) or []
        errors = payload.get("errors", []) or []

        def fmt_list(items):
            return "\n".join(f"- {x}" for x in items) if items else "(none)"

        # Variables tab: ONLY local design variables
        variables_text = "Local design variables (oDesign.GetVariables equivalent):\n" + fmt_list(design_vars)

        # Outputs tab: ONLY report names
        outputs_text = "HFSS report names (ReportSetup.GetAllReportNames equivalent):\n" + fmt_list(report_names)

        if errors:
            outputs_text += "\n\n---- Notes / Errors ----\n" + "\n".join(f"- {e}" for e in errors)

        dlg = MetadataDialog(variables_text=variables_text, outputs_text=outputs_text, parent=self)
        dlg.exec_()


def main():
    project_path = sys.argv[1] if len(sys.argv) > 1 else None

    if project_path:
        print("Received AEDT project path:", project_path)

    app = QApplication(sys.argv)
    win = MainWindow(project_path=project_path)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()