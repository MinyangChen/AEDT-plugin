import sys
import os
import random

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTextEdit,
    QPushButton,
    QMessageBox,
    QDialog,
    QTabWidget,
    QLabel,
    QLineEdit,
)
from PyQt5.QtGui import QIntValidator
from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal


# -------------------- Random parameter configuration --------------------
# Default seed shown in the GUI.
PARAMETER_SEED = 2025

# These variable names must match the HFSS local design-variable names exactly.
PARAMETER_SPECS = {
    "R1": {"range": [4, 7], "unit": "mm"},
    "H1": {"range": [0, 1], "unit": "mm"},
    "fill_gap": {"range": [0, 1], "unit": "mm"},
    "helixz": {"range": [0, 1], "unit": "mm"},
    "helixx": {"range": [0.1, 0.5], "unit": "mm"},
    "nturn": {"range": [0.8, 1], "unit": ""},
    "hpitch": {"range": [0, 1], "unit": "mm"},
    "H_Feed": {"range": [1.5, 3], "unit": "mm"},
    "Probe_length_in": {"range": [0, 1], "unit": "mm"},
    "feed_angle": {"range": [-4, 2.55], "unit": "deg"},
    "Probe_Cap_R": {"range": [0, 1], "unit": "mm"},
    "Probe_Cap_H": {"range": [0.1, 3], "unit": "mm"},
}


# -------------------- Popup dialogs --------------------

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


class ParametersDialog(QDialog):
    """Popup used to show the parameter values that were just generated/applied."""

    def __init__(self, parameters_text: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Applied Parameters")
        self.resize(700, 500)

        text_view = QTextEdit()
        text_view.setReadOnly(True)
        text_view.setText(parameters_text)

        layout = QVBoxLayout()
        layout.addWidget(text_view)
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

        # Stores the last parameter set that was successfully applied through the GUI.
        # Format: {var_name: "value_with_unit"}
        self.applied_parameters = None
        self.last_applied_seed = None

        self.init_ui()

    # ---------- GUI setup ----------

    def init_ui(self):
        title = "Optimizer GUI"
        if self.project_path:
            title += f" - {self.project_path}"
        self.setWindowTitle(title)

        self.resize(650, 450)
        self.setAttribute(Qt.WA_DeleteOnClose, True)

        layout = QVBoxLayout()

        # Seed input row
        seed_row = QHBoxLayout()
        seed_label = QLabel("Random Seed:")
        self.seed_input = QLineEdit(str(PARAMETER_SEED))
        self.seed_input.setValidator(QIntValidator(0, 2147483647, self))
        self.seed_input.setToolTip("Change the seed, then click Apply Parameters.")
        seed_row.addWidget(seed_label)
        seed_row.addWidget(self.seed_input)
        layout.addLayout(seed_row)

        self.apply_params_btn = QPushButton("Apply Parameters")
        self.apply_params_btn.clicked.connect(self.apply_parameters)
        layout.addWidget(self.apply_params_btn)

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

    def _release_aedt_connection(self):
        """Release the current PyAEDT desktop connection cleanly."""
        if self.desktop:
            try:
                self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                self.append_log("Released PyAEDT Desktop connection.")
            except Exception as e:
                self.append_log(f"WARNING: Failed to release Desktop cleanly: {e}")
        self.desktop = None
        self.app = None

    # ---------- Parameter helpers ----------

    def _get_seed_value(self):
        """Read the seed from the GUI."""
        text = self.seed_input.text().strip()
        if not text:
            raise ValueError("Please enter a random seed.")
        return int(text)

    def _format_hfss_expression(self, value: float, unit: str) -> str:
        """Format a numeric value as an HFSS-compatible string expression."""
        value_text = f"{value:.8f}".rstrip("0").rstrip(".")
        return f"{value_text}{unit}" if unit else value_text

    def _generate_random_parameter_expressions(self, seed: int):
        """
        Generate reproducible random parameter values from the configured ranges.
        Returns a dict like: {"R1": "5.123mm", "feed_angle": "1.02deg", ...}
        """
        rng = random.Random(seed)
        generated = {}

        for var_name, spec in PARAMETER_SPECS.items():
            low, high = spec["range"]
            value = rng.uniform(low, high)
            generated[var_name] = self._format_hfss_expression(value, spec["unit"])

        return generated

    def _apply_parameter_expressions_to_hfss(self, parameter_expressions: dict):
        """
        Apply a parameter dictionary to the active HFSS design.
        """
        if not self.app:
            raise RuntimeError("PyAEDT app not available.")

        design_var_names = set(self.app.variable_manager.design_variable_names or [])
        missing = [name for name in parameter_expressions if name not in design_var_names]
        if missing:
            raise RuntimeError(
                "These configured variables were not found in the active HFSS design:\n"
                + "\n".join(missing)
            )

        for var_name, expression in parameter_expressions.items():
            ok = self.app.variable_manager.set_variable(
                name=var_name,
                expression=expression,
                overwrite=True,
            )
            if not ok:
                raise RuntimeError(f"Failed to set variable '{var_name}' to '{expression}'.")

    def _format_parameters_for_popup(self, parameter_expressions: dict, seed: int) -> str:
        lines = [
            "Applied random parameter set",
            f"Seed = {seed}",
            "",
        ]
        for var_name, expression in parameter_expressions.items():
            lines.append(f"{var_name} = {expression}")
        return "\n".join(lines)

    # ---------- GUI action: apply parameters ----------

    def apply_parameters(self):
        """
        Generate a reproducible random parameter set from the current seed,
        write it to HFSS, store it, and show it in a popup.
        """
        self.apply_params_btn.setEnabled(False)
        try:
            seed = self._get_seed_value()

            self.append_log("Attempting to connect to open AEDT session for parameter update...")
            if not self.connect_to_open_aedt():
                QMessageBox.critical(
                    self,
                    "Parameter Update Failed",
                    "Failed to connect to AEDT.\nCheck the log for details."
                )
                return

            parameter_expressions = self._generate_random_parameter_expressions(seed)

            self.append_log(f"Applying parameter set generated from seed {seed}...")
            self._apply_parameter_expressions_to_hfss(parameter_expressions)

            self.applied_parameters = parameter_expressions
            self.last_applied_seed = seed

            for var_name, expression in parameter_expressions.items():
                self.append_log(f"Applied: {var_name} = {expression}")

            dlg = ParametersDialog(
                parameters_text=self._format_parameters_for_popup(parameter_expressions, seed),
                parent=self,
            )
            dlg.exec_()

        except Exception as e:
            self.append_log(f"ERROR while applying parameters: {e}")
            QMessageBox.critical(self, "Parameter Update Failed", str(e))

        finally:
            self._release_aedt_connection()
            self.apply_params_btn.setEnabled(True)

    def run_hfss_simulation_and_get_s11_summary(self):
        """
        Run HFSS and try to read S(1,1).
        Returns:
            (summary_text, error_text)
        Exactly one of them is usually None.
        """
        if not self.app:
            return None, "PyAEDT app is not available."

        try:
            self.append_log("Starting HFSS simulation via app.analyze()...")
            ok = self.app.analyze()
            self.append_log(f"HFSS analyze() returned: {ok}")

            # Protection 1:
            # If solve failed, stop here instead of trying to read solution data.
            if not ok:
                self.append_log("WARNING: HFSS solve failed or no valid solution was produced.")
                return None, (
                    "HFSS simulation failed or no valid solution was produced.\n"
                    "Please try another seed and check the AEDT Message Manager."
                )

        except Exception as e:
            self.append_log(f"ERROR during analyze(): {e}")
            return None, f"HFSS simulation failed during analyze().\n{e}"

        try:
            self.append_log("Retrieving dB(S(1,1)) solution data...")

            setup_sweep = self.app.nominal_sweep
            data = self.app.post.get_solution_data(
                expressions="S(1,1)",
                setup_sweep_name=setup_sweep,
            )

            # Protection 2:
            # get_solution_data may return False/None or another invalid object.
            if not data or not hasattr(data, "get_expression_data"):
                self.append_log("WARNING: No valid solution-data object was returned.")
                return None, (
                    "HFSS simulation completed, but no valid S(1,1) solution data was returned.\n"
                    "Please try another seed and check the AEDT Message Manager."
                )

            freqs, s11_db = data.get_expression_data(
                expression="S(1,1)",
                formula="db20",
            )

            n = len(freqs)
            if n == 0:
                self.append_log("No data points returned from get_expression_data().")
                return None, (
                    "HFSS simulation completed, but S(1,1) returned no data points.\n"
                    "Please try another seed and check the AEDT Message Manager."
                )

            lines = [f"Retrieved {n} points of dB(S(1,1)) (showing first few):"]
            max_to_show = min(5, n)
            for i in range(max_to_show):
                lines.append(f"{i}: Freq = {freqs[i]}, dB(S11) = {s11_db[i]}")
            return "\n".join(lines), None

        except Exception as e:
            self.append_log(f"ERROR retrieving solution data: {e}")
            return None, (
                "HFSS simulation finished, but reading S(1,1) failed.\n"
                f"{e}"
            )

    # ---------- Simulation entry ----------

    def run_simulation(self):
        self.start_btn.setEnabled(False)
        try:
            if not self.applied_parameters:
                QMessageBox.warning(
                    self,
                    "No Parameters Applied",
                    "Please click 'Apply Parameters' first."
                )
                return

            self.append_log("Attempting to connect to open AEDT session...")
            if not self.connect_to_open_aedt():
                self.append_log("Aborting simulation because connection failed.")
                QMessageBox.critical(
                    self,
                    "Simulation Failed",
                    "Failed to connect to AEDT.\nCheck the log for details."
                )
                return

            # Re-apply the last parameter set so the solve always uses
            # the parameters most recently sent from the GUI.
            if self.last_applied_seed is not None:
                self.append_log(
                    f"Re-applying the last parameter set before simulation (seed={self.last_applied_seed})..."
                )
            else:
                self.append_log("Re-applying the last parameter set before simulation...")

            self._apply_parameter_expressions_to_hfss(self.applied_parameters)

            summary, error_message = self.run_hfss_simulation_and_get_s11_summary()

            if summary:
                QMessageBox.information(self, "Simulation Result", summary)
            else:
                QMessageBox.warning(
                    self,
                    "Simulation Failed",
                    error_message or "Simulation failed. Check the log for details."
                )

        except Exception as e:
            self.append_log(f"ERROR during simulation flow: {e}")
            QMessageBox.critical(self, "Simulation Failed", str(e))

        finally:
            self._release_aedt_connection()
            self.start_btn.setEnabled(True)

    # ---------- Fetch Metadata ----------

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

        variables_text = "Local design variables (oDesign.GetVariables equivalent):\n" + fmt_list(design_vars)
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