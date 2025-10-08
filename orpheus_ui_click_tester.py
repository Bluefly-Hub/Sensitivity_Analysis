import uiautomation as uia
import subprocess, sys
import time
import pythoncom
from pathlib import Path
from pywinauto import Application, Desktop
from pywinauto.uia_defines import NoPatternInterfaceError

# Make misses fast
uia.SetGlobalSearchTimeout(0.3)


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

def Sensitivity_Analysis2():
    """ Click Sensitivity Analysis menu item in Orpheus main window """
    # 1) Attach to your Orpheus window (adjust RegexName if needed)
    Main_Window = uia.WindowControl(RegexName=r'Orpheus*')

    # Many WPF/WinForms apps keep menu items in the tree even when the submenu isn't dropped.
    target = uia.MenuItemControl(searchFromControl=Main_Window, Name='Sensitivity Analysis...')
    inv = target.GetInvokePattern()
    #try:
    inv.Invoke()
    inv = None





def exit():
    # Attach to the main window
    Main_Window = uia.WindowControl(RegexName=r'Orpheus*', Depth=1)

    # Drill directly to the Exit button by AutomationId
    exit_btn = Main_Window.ButtonControl(AutomationId="cmdExit")

    # Invoke the button
    exit_btn.GetInvokePattern().Invoke()

if __name__ == "__main__":
    Sensitivity_Analysis()
    exit()
