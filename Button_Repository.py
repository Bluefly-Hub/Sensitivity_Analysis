import subprocess, sys
from pathlib import Path
from cerberus_sensitivity.automation.button_bridge_CSharp_to_py import invoke_button, list_buttons, set_button_value

# def Sensitivity_Analysis():
#     """Invoke the Sensitivity Analysis menu via the compiled C# helper."""
#     ...
# Keeping the layout similar to your original script, but pruning unused imports.


def Sensitivity_Analysis():
    """Launch the helper executable to open the Sensitivity Analysis menu."""
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
    """Navigate to Tools > Sensitivity Analysis... via repository button1."""
    return invoke_button("button1", timeout=timeout)


def button_exit_wizard(timeout: float = 90.0):
    """Invoke the wizard Exit button (button2)."""
    return invoke_button("button2", timeout=timeout)

def File_OpenTemplate(timeout: float = 90.0):
    """Open Sensitivity Analysis > File > Open Template (button3)."""
    return invoke_button("button3", timeout=timeout)

def File_OpenTemplate_auto(timeout: float = 90.0):
    """Choose the 'auto' template entry (button4)."""
    return invoke_button("button4", timeout=timeout)

def Parameters_Pipe_fluid_density(timeout: float = 90.0):
    """Toggle the Pipe Fluid Density parameter checkbox (button5)."""
    return invoke_button("button5", timeout=timeout)

def Parameters_POOH(timeout: float = 90.0):
    """Toggle the POOH parameter checkbox (button6)."""
    return invoke_button("button6", timeout=timeout)

def Sensitivity_Setting_Outputs(timeout: float = 90.0):
    """Select the Outputs pane tab (button7)."""
    return invoke_button("button7", timeout=timeout)

def Parameters_Maximum_Surface_Weight_During_POOH(timeout: float = 90.0):
    """Toggle Maximum surface weight during POOH (button8)."""
    return invoke_button("button8", timeout=timeout)

def Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(timeout: float = 90.0):
    """Toggle Maximum pipe stress during POOH (% of YS) (button9)."""
    return invoke_button("button9", timeout=timeout)

def Parameter_Matrix_Wizard(timeout: float = 90.0):
    """Open the Parameter Matrix Wizard button (button10)."""
    return invoke_button("button10", timeout=timeout)

def Parameter_Matrix_BHA_Depth_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for BHA Depth row 0 (button11)."""
    return invoke_button("button11", timeout=timeout)

def Value_List_Depth_1(timeout: float = 60.0):
    """Open the value list entry for Depth value 1 (button12)."""
    return invoke_button("button12", timeout=timeout)

def Edit_cmdDelete(timeout: float = 60.0):
    """Trigger the Delete button in the value list editor (button13)."""
    return invoke_button("button13", timeout=timeout)

def Parameter_Value_Editor_Set_Value(val: str, timeout: float = 10):
    """Set the depth value in the editor (button14)."""
    return set_button_value("button14", value=val, timeout=timeout)

def Edit_cmdAdd(timeout: float = 60.0):
    """Trigger the Add button in the value list editor (button15)."""
    return invoke_button("button15", timeout=timeout)

def Edit_cmdOK(timeout: float = 60.0):
    """Trigger the OK button in the value list editor (button16)."""
    return invoke_button("button16", timeout=timeout)

def Parameter_Matrix_PFD_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for PFD row 0 (button17)."""
    return invoke_button("button17", timeout=timeout)

def Value_List_PFD_1(timeout: float = 60.0):
    """Open the value list entry for PFD value 1 (button18)."""
    return invoke_button("button18", timeout=timeout)

def Parameter_Matrix_FOE_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for FOE row 0 (button19)."""
    return invoke_button("button19", timeout=timeout)


def Value_List_FOE_1(timeout: float = 60.0):
    """Open the value list entry for FOE value 1 (button20)."""
    return invoke_button("button20", timeout=timeout)

def Value_List_Item0(timeout: float = 60.0):
    """Open the value list entry for FOE value 2 (button21)."""
    return invoke_button("button21", timeout=timeout)

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
    # Parameter_Matrix_BHA_Depth_Row0()
    # Value_List_Depth_1()
    # Edit_cmdDelete()
    # Parameter_Value_Editor_Set_Value("1000")
    # Edit_cmdAdd()
    # Edit_cmdOK() 
    # Parameter_Matrix_PFD_Row0()
    # Value_List_PFD_1()
    # Parameter_Matrix_FOE_Row0()
    # Value_List_FOE_1()
    Value_List_Item0()


   # list_buttons


