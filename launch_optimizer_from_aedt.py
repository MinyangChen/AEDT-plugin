# launch_optimizer_from_aedt.py
# Custom PyAEDT extension:
# - attaches to current AEDT session
# - gets the active AEDT project path
# - starts optimizer_gui.py and passes the path as an argument

import os
import subprocess
from pathlib import Path
import ansys.aedt.core  # PyAEDT

# --- Attach to the running AEDT session ---
if "PYAEDT_SCRIPT_PORT" in os.environ and "PYAEDT_SCRIPT_VERSION" in os.environ:
    port = os.environ["PYAEDT_SCRIPT_PORT"]
    version = os.environ["PYAEDT_SCRIPT_VERSION"]
else:
    port = 0
    version = "2025.2"

desktop = ansys.aedt.core.Desktop(
    new_desktop_session=False,
    specified_version=version,
    port=port,
)

log = desktop.logger  # PyAEDT AedtLogger

# --- Get active project path ---
try:
    project_path = desktop.project_path()  # path to current .aedt
    if project_path:
        log.info("Active AEDT project path: %s", project_path)
    else:
        log.info("Active AEDT project path: <none / project not saved yet>")
except Exception as e:
    log.info("Failed to get AEDT project path: %s", e)
    project_path = ""

# --- Configure paths (EDIT ONLY THESE TWO LINES IF NEEDED) ---

PYTHON_EXE = r"C:\Program Files\ANSYS Inc\v252\commonfiles\CPython\3_10\winx64\Release\python\python.exe"
GUI_SCRIPT = str(Path(__file__).resolve().parent / "optimizer_gui.py")

# PYTHON_EXE = r"C:\Users\2948856C\OneDrive - University of Glasgow\0_PHD\2_software_projects\2_pysadeaGUI\ai-dad-gui\sadeagui_env\Scripts\python.exe"
# GUI_SCRIPT = r"C:\Users\2948856C\OneDrive - University of Glasgow\0_PHD\2_software_projects\9_AllGUI\ai-dad-gui\aide_gui_run.py"


# --- Check script exists ---
gui_path = Path(GUI_SCRIPT)
if not gui_path.exists():
    log.info("optimizer_gui.py not found at: %s", gui_path)
    raise FileNotFoundError(f"optimizer_gui.py not found at: {gui_path}")

# --- Build command to launch GUI and pass project_path ---
cmd = [PYTHON_EXE, str(gui_path)]
if project_path:
    cmd.append(project_path)  # argv[1] in optimizer_gui.py

# --- Launch GUI (non-blocking) ---
try:
    p = subprocess.Popen(cmd)
    log.info("SADEA_GUI launched, PID=%s", p.pid)
except Exception as e:
    log.info("Failed to launch Optimizer GUI: %s", e)

# --- Detach from AEDT (don't close it) ---
desktop.release_desktop(False, False)
