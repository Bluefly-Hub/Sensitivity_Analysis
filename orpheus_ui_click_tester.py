import uiautomation as uia
import subprocess, sys
import time
import pythoncom
from pathlib import Path
from pywinauto import Application, Desktop
from pywinauto.uia_defines import NoPatternInterfaceError
from cerberus_sensitivity.automation.button_bridge import invoke_button

# Make misses fast
uia.SetGlobalSearchTimeout(1)


def Sensitivity_Analysis():
    """Invoke the Sensitivity Analysis menu via the compiled C# helper."""
    exe_path = Path(__file__).resolve().parent / "bin" / "Debug" / "net9.0-windows" / "Test_C.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"Expected helper executable not found: {exe_path}")

    try:
        result = subprocess.run(
            [str(exe_path)],
            check=True,
            capture_output=True,
            text=True,
        )
    except subprocess.CalledProcessError as exc:
        if exc.stdout:
            print(exc.stdout, end="", file=sys.stdout)
        if exc.stderr:
            print(exc.stderr, end="", file=sys.stderr)
        raise RuntimeError(f"Sensitivity_Analysis failed via helper executable: {exe_path}") from exc
    else:
        if result.stdout:
            print(result.stdout, end="")
        if result.stderr:
            print(result.stderr, end="", file=sys.stderr)


def button_open_template(timeout: float = 180.0):
    """Fire the repository button that opens Sensitivity Analysis -> File -> Open Template."""
    return invoke_button("button1", timeout=timeout)


def button_exit_wizard(timeout: float = 90.0):
    """Fire the repository button that closes the Sensitivity Analysis wizard."""
    return invoke_button("button2", timeout=timeout)

def File_OpenTemplate(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template."""
    return invoke_button("button3", timeout=timeout)

def File_OpenTemplate_auto(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button4", timeout=timeout)

if __name__ == "__main__":
    #button_open_template()
    #File_OpenTemplate()
    File_OpenTemplate_auto()

