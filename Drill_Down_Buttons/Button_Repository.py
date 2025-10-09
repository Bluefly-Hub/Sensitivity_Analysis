import uiautomation as uia
import subprocess, sys
import time
import pythoncom
from pathlib import Path
from pywinauto import Application, Desktop
from pywinauto.uia_defines import NoPatternInterfaceError
from cerberus_sensitivity.automation.button_bridge_CSharp_to_py import invoke_button, set_button_value

# Make misses fast
uia.SetGlobalSearchTimeout(1)


def Sensitivity_Analysis():
    """Invoke the Sensitivity Analysis menu via the compiled C# helper."""
    exe_path = Path(__file__).resolve().parent / "bin" / "Debug" / "net9.0-windows" / "Drill_Down_With_C.exe"
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

def Parameters_Pipe_fluid_density(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button5", timeout=timeout)

def Parameters_POOH(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button6", timeout=timeout)

def Sensitivity_Setting_Outputs(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button7", timeout=timeout)

def Parameters_Maximum_Surface_Weight_During_POOH(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button8", timeout=timeout)

def Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button9", timeout=timeout)

def Parameter_Matrix_Wizard(timeout: float = 90.0):
    """Fire the repository button that opens the Sensitivity Analysis -> File -> Open Template->Auto."""
    return invoke_button("button10", timeout=timeout)

def Parameter_Matrix_BHA_Depth_Row0(value: str, timeout: float = 90.0):
    """Set the Parameter Matrix row 0 BHA Depth value."""
    return set_button_value("button11", value, timeout=timeout)

if __name__ == "__main__":
    # button_open_template()
    # File_OpenTemplate()
    # File_OpenTemplate_auto()
    # Parameters_Pipe_fluid_density()
    # Parameters_POOH()
    # Sensitivity_Setting_Outputs()
    # Parameters_Maximum_Surface_Weight_During_POOH()
    # Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS()
    # Parameter_Matrix_Wizard()
    Parameter_Matrix_BHA_Depth_Row0("5000")

