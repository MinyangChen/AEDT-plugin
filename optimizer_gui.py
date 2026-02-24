# optimizer_gui.py
# Minimal PyQt5 GUI that:
#   - receives the AEDT project path (optional, for display)
#   - attaches to the already-open AEDT via PyAEDT
#   - runs an HFSS simulation on the active design
#   - reads dB(S(1,1)) vs Freq and shows a summary in the GUI
#
# Requires: PyQt5 + PyAEDT installed in this Python interpreter.

import sys
import os

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTextEdit
from PyQt5.QtCore import Qt, QTimer


class MainWindow(QWidget):
    def __init__(self, project_path=None, parent=None):
        super().__init__(parent)
        self.project_path = project_path
        self.desktop = None
        self.app = None  # PyAEDT app (HFSS, etc.)

        self.init_ui()

        # Start the simulation once the GUI event loop is running.
        QTimer.singleShot(0, self.run_simulation)

    # ---------- GUI setup ----------

    def init_ui(self):
        title = "Optimizer GUI"
        if self.project_path:
            title += f" - {self.project_path}"
        self.setWindowTitle(title)

        self.resize(600, 400)
        self.setAttribute(Qt.WA_DeleteOnClose, True)

        layout = QVBoxLayout()
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.setLayout(layout)

        self.append_log("Optimizer GUI started.")
        if self.project_path:
            self.append_log(f"Received AEDT project path: {self.project_path}")

    def append_log(self, text: str):
        """Append a line of text to the log box."""
        self.log.append(text)

    # ---------- AEDT / PyAEDT helpers ----------

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
            # Attach to existing AEDT (do NOT start a new one) 
            self.desktop = Desktop(
                version=version,
                new_desktop=False,
                port=int(port),
            )

            # Get the active design as a PyAEDT app (HFSS, etc.) 
            from ansys.aedt.core.generic.design_types import get_pyaedt_app
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

    def run_hfss_simulation_and_get_s11(self):
        """Run HFSS analyze() and read dB(S(1,1)) vs Freq using PyAEDT."""
        if not self.app:
            self.append_log("ERROR: PyAEDT app not available.")
            return

        # 1) Run the HFSS simulation (blocking until it finishes)
        try:
            self.append_log("Starting HFSS simulation via app.analyze()...")
            ok = self.app.analyze()  # solve all setups of the active design
            self.append_log(f"HFSS analyze() returned: {ok}")
            if not ok:
                self.append_log("WARNING: analyze() returned False (simulation may have failed).")
        except Exception as e:
            self.append_log(f"ERROR during analyze(): {e}")
            return

        # 2) Get dB(S(1,1)) vs Freq from the solved setup
        try:
            self.append_log("Retrieving dB(S(1,1)) solution data...")

            # Name of the nominal sweep for the active design
            setup_sweep = self.app.nominal_sweep  # e.g. "Setup1 : LastAdaptive" 

            # Get SolutionData object for S(1,1) 
            data = self.app.post.get_solution_data(
                expressions="S(1,1)",
                setup_sweep_name=setup_sweep,
            )

            # New API: get_expression_data returns (x, y) arrays 
            freqs, s11_db = data.get_expression_data(
                expression="S(1,1)",
                formula="db20",          # get dB20(S(1,1)); change to "db10" if preferred
            )

            n = len(freqs)
            if n == 0:
                self.append_log("No data points returned from get_expression_data().")
                return

            self.append_log(f"Retrieved {n} points of dB(S(1,1)). Showing first few:")
            max_to_show = min(5, n)
            for i in range(max_to_show):
                self.append_log(f"  {i}: Freq = {freqs[i]}, dB(S11) = {s11_db[i]}")

        except Exception as e:
            self.append_log(f"ERROR retrieving solution data: {e}")


    # ---------- Main simulation entry ----------

    def run_simulation(self):
        """High-level: connect to AEDT, run simulation, pull results, then release."""
        self.append_log("Attempting to connect to open AEDT session...")
        if not self.connect_to_open_aedt():
            self.append_log("Aborting simulation because connection failed.")
            return

        self.run_hfss_simulation_and_get_s11()

        # Release the Desktop handle when done (AEDT remains open) 
        if self.desktop:
            try:
                self.desktop.release_desktop(close_projects=False, close_on_exit=False)
                self.append_log("Released PyAEDT Desktop connection.")
            except Exception as e:
                self.append_log(f"WARNING: Failed to release Desktop cleanly: {e}")


def main():
    # Receive AEDT project path from the launcher (for display only)
    project_path = sys.argv[1] if len(sys.argv) > 1 else None

    if project_path:
        print("Received AEDT project path:", project_path)

    app = QApplication(sys.argv)
    win = MainWindow(project_path=project_path)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
