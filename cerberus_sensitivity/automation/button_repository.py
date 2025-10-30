from __future__ import annotations

import re
import time
from ctypes import windll, wintypes
from pathlib import Path
from typing import Any, Mapping, Optional, Sequence
import warnings

import pandas as pd

from pywinauto import Application, timings
from pywinauto.findbestmatch import MatchError
from pywinauto.findwindows import ElementAmbiguousError, ElementNotFoundError
from pywinauto.remote_memory_block import AccessDenied
from concurrent.futures import ThreadPoolExecutor, TimeoutError as InvokeTimeout

warnings.filterwarnings(
    "ignore",
    message="32-bit application should be automated using 32-bit Python",
    category=UserWarning,
    module="pywinauto.application",
)

timings.Timings.window_find_timeout = 1.0
timings.Timings.window_find_retry = 0.05
timings.Timings.after_click_wait = 0.0

_FIND_WINDOW_EX = windll.user32.FindWindowExW
_FIND_WINDOW_EX.argtypes = [
    wintypes.HWND,
    wintypes.HWND,
    wintypes.LPCWSTR,
    wintypes.LPCWSTR,
]
_FIND_WINDOW_EX.restype = wintypes.HWND

_IS_WINDOW = windll.user32.IsWindow

_APP_UIA: Optional[Application] = None
_APP_WIN32: Optional[Application] = None
_MAIN_HANDLE: Optional[int] = None
_VALUE_LIST_HANDLE: Optional[int] = None

from .button_bridge_CSharp_to_py import collect_table, invoke_button, list_buttons, set_button_value

_DEFAULT_DUMP_PATH = (
    Path(__file__).resolve().parents[2] / "inspect_dumps" / "Windows_Inspect_Dump.txt"
)


def _table_to_dataframe(
    table_payload: Mapping[str, Any],
    *,
    headers_override: Sequence[str] | None = None,
) -> pd.DataFrame:
    rows = list(table_payload.get("Rows") or [])
    headers = list(headers_override or table_payload.get("Headers") or [])

    if rows and headers and len(headers) != len(rows[0]):
        if len(headers) < len(rows[0]):
            headers.extend(f"Column {i}" for i in range(len(headers), len(rows[0])))
        else:
            headers = headers[: len(rows[0])]

    columns = headers if headers else None
    df = pd.DataFrame(rows, columns=columns)
    return df


def _set_checkbox_state(button_key: str, checked: bool, *, timeout: float = 90.0):
    value = "True" if checked else "False"
    completed = set_button_value(button_key, value=value, timeout=timeout)
    stdout = (completed.stdout or "").splitlines()
    for line in stdout:
        if "Toggle.ToggleState" in line:
            return line.strip()
    combined = "\n".join(line.strip() for line in stdout if line.strip())
    if combined:
        return combined
    state = "On (1)" if checked else "Off (0)"
    return f"Toggle.ToggleState:\t{state}"


def Sensitivity_Setting_Outputs(timeout: float = 90.0):
    """Select the Outputs pane tab (button7)."""
    return invoke_button("button7", timeout=timeout)


def Sensitivity_Setting_Parameters(timeout: float = 90.0):
    """Select the Parameters pane tab (button27)."""
    return invoke_button("button27", timeout=timeout)


def _ensure_parameters_tab(timeout: float = 90.0) -> None:
    Sensitivity_Setting_Parameters(timeout=timeout)
    #time.sleep(0.1)


def _ensure_outputs_tab(timeout: float = 90.0) -> None:
    Sensitivity_Setting_Outputs(timeout=timeout)
    #time.sleep(0.1)


def _wait_for_parameter_matrix_window(
    *,
    timeout: float = 15.0,
    retry_interval: float = 0.1,
) -> None:
    """Ensure the Parameter Matrix window is visible before interacting."""
    deadline = time.perf_counter() + timeout
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_window = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    while time.perf_counter() < deadline:
        try:
            matrix_window = main_window.child_window(
                auto_id="frmSensitivityMatrix",
            )
            matrix_window.wait("visible ready", timeout=0.3)
            try:
                matrix_window.set_focus()
            except (RuntimeError, timings.TimeoutError):
                pass
            return
        except (ElementNotFoundError, ElementAmbiguousError, MatchError, timings.TimeoutError):
            time.sleep(retry_interval)

    raise TimeoutError("Timed out waiting for Parameter Matrix window to become ready.")


def _coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    numeric_pattern = r"([-+]?\d*\.?\d+)"

    for column in df.columns:
        extracted = df[column].astype(str).str.extract(numeric_pattern)[0]
        if extracted.notna().any():
            numeric = pd.to_numeric(extracted, errors="coerce")
            if numeric.notna().any():
                df[column] = numeric
    return df


def _headers_from_dump(button_key: str, dump_path: Path | None = None) -> list[str]:
    path = dump_path or _DEFAULT_DUMP_PATH
    try:
        text = Path(path).read_text(encoding="utf-8", errors="ignore")
    except FileNotFoundError:
        return []

    section_pattern = re.compile(rf"\[{re.escape(button_key)}\](.*?)(?:\n\[|$)", re.DOTALL)
    match = section_pattern.search(text)
    if not match:
        return []

    header_block_match = re.search(
        r"Table\.ColumnHeaders:(.*?)(?:\n[A-Z][^\n]*:|\Z)", match.group(1), re.DOTALL
    )
    raw_headers: list[str] = []
    if header_block_match:
        raw_headers = re.findall(r'"([^"]+)"', header_block_match.group(1))

    if not raw_headers:
        children_match = re.search(
            r"Children:(.*?)(?:\n[A-Z][^\n]*:|\Z)", match.group(1), re.DOTALL
        )
        if children_match:
            pending: list[str] = []
            for line in children_match.group(1).splitlines():
                stripped = line.strip()
                if not stripped:
                    continue

                if stripped.endswith('" header'):
                    core = stripped[:-8].strip('"')
                    if pending:
                        pending.append(core)
                        combined = " ".join(part for part in pending if part)
                        raw_headers.append(combined)
                        pending = []
                    else:
                        raw_headers.append(core)
                else:
                    pending.append(stripped.strip('"'))

            if pending:
                combined = " ".join(part for part in pending if part)
                if combined:
                    raw_headers.append(combined)

    return [" ".join(value.split()) for value in raw_headers if value.strip()]


def button_Sensitivity_Analysis(timeout: float = 180.0):
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


def Parameters_Pipe_fluid_density(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set the Pipe Fluid Density parameter checkbox (button5)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button5", timeout=timeout)
    return _set_checkbox_state("button5", checked, timeout=timeout)

def Parameters_Depth(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set the Pipe Depth parameter checkbox (button5)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button32", timeout=timeout)
    return _set_checkbox_state("button32", checked, timeout=timeout)

def Parameters_FF(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set the Pipe Depth parameter checkbox (button5)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button33", timeout=timeout)
    return _set_checkbox_state("button33", checked, timeout=timeout)

def Parameters_POOH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set the POOH parameter checkbox (button6)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button6", timeout=timeout)
    return _set_checkbox_state("button6", checked, timeout=timeout)


def Setup_POOH(timeout: float = 90.0):
    """Ensure the application is configured for POOH batches."""
    Parameters_POOH(checked=True, timeout=timeout)
    Parameters_RIH(checked=False, timeout=timeout)
    Parameters_Maximum_Surface_Weight_During_POOH(checked=True, timeout=timeout)
    Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(checked=True, timeout=timeout)
    Parameters_Minimum_Surface_Weight_During_RIH(checked=False, timeout=timeout)



def Parameters_Maximum_Surface_Weight_During_POOH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set Maximum surface weight during POOH (button8)."""
    _ensure_outputs_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button8", timeout=timeout)
    return _set_checkbox_state("button8", checked, timeout=timeout)


def Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or set Maximum pipe stress during POOH (% of YS) (button9)."""
    _ensure_outputs_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button9", timeout=timeout)
    return _set_checkbox_state("button9", checked, timeout=timeout)


def Parameter_Matrix_Wizard(timeout: float = 90.0):
    """Open the Parameter Matrix Wizard button (button10)."""
    return invoke_button("button10", timeout=timeout)


def Parameter_Matrix_BHA_Depth_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for BHA Depth row 0 (button11)."""
    _wait_for_parameter_matrix_window(timeout=min(timeout, 15.0))
    return invoke_button("button11", timeout=timeout)


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
    _wait_for_parameter_matrix_window(timeout=min(timeout, 15.0))
    return invoke_button("button17", timeout=timeout)


def Parameter_Matrix_FOE_POOH_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for FOE row 0 (button19)."""
    _wait_for_parameter_matrix_window(timeout=min(timeout, 15.0))
    return invoke_button("button19", timeout=timeout)

def Parameter_Matrix_FOE_RIH_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for FOE row 0 (button28)."""
    _wait_for_parameter_matrix_window(timeout=min(timeout, 15.0))
    return invoke_button("button28", timeout=timeout)


def Value_List_Item0(timeout: float = 60.0):
    """Open the value list entry for FOE value 2 (button21)."""
    return invoke_button("button21", timeout=timeout)


def Sensitivity_Analysis_Calculate(timeout: float = 60.0):
    """Trigger the PAD Calculate button (button22)."""
    return invoke_button("button22", timeout=timeout)


def Parameters_Minimum_Surface_Weight_During_RIH(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or set Minimum surface weight during RIH (button23)."""
    _ensure_outputs_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button23", timeout=timeout)
    return _set_checkbox_state("button23", checked, timeout=timeout)


def Sensitivity_Table(timeout: float = 90.0):
    """Sensitivity Table (button24)."""
    table_data = collect_table("button24", timeout=timeout)
    df = _table_to_dataframe(table_data)
    if df.columns.tolist() and df.columns[0] == "#":
        df = df.drop(columns="#")

    df = df.loc[:, df.apply(lambda col: col.astype(str).str.strip().ne("").any())]

    dump_headers = _headers_from_dump("button25") or _headers_from_dump("button24")
    if dump_headers:
        cleaned_headers = [header for header in dump_headers if header != "#"]
        target_headers = cleaned_headers[: len(df.columns)]
        if len(target_headers) == len(df.columns):
            rename_map = dict(zip(df.columns, target_headers))
            df = df.rename(columns=rename_map)

    if df.columns.tolist() and df.columns[0] == "#":
        df = df.drop(columns="#")

    df = df.loc[:, df.apply(lambda col: col.astype(str).str.strip().ne("").any())]

    df = _coerce_numeric_columns(df)
    return df


def Parameters_RIH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set the RIH parameter checkbox (button26)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button26", timeout=timeout)
    return _set_checkbox_state("button26", checked, timeout=timeout)

def Sensitivity_Parameter_ok(timeout: float = 60.0):
    """Trigger the PAD OK button (button29)."""
    return invoke_button("button29", timeout=timeout)

def list_defined_buttons(*, timeout: float = 30.0):
    return list_buttons(timeout=timeout)

def invoke_with_timeout(ctrl, timeout=1.0):
    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(ctrl.invoke)
        try:
            return future.result(timeout=timeout)
        except InvokeTimeout:
            raise TimeoutError(f"invoke() timed out after {timeout} s")



def find_child_by_class(parent: int, class_name: str, automation_id: str, app_uia: Application) -> int:
    child = _FIND_WINDOW_EX(parent, 0, class_name, None)
    while child:
        if app_uia.window(handle=child).element_info.automation_id == automation_id:
            return child
        child = _FIND_WINDOW_EX(parent, child, class_name, None)
    raise RuntimeError(f"Unable to locate {automation_id}")


def Populate_Value_List(values: Sequence[str]) -> None:
    start = time.perf_counter()

    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    app_win32 = Application(backend="win32").connect(process=app_uia.process)
    main_win32 = app_win32.window(handle=main_uia.handle)

    pnl_numeric = main_uia.child_window(auto_id="pnlNumeric", control_type="Pane")
    panel_handle = pnl_numeric.handle

    txt_handle = find_child_by_class(panel_handle, "WindowsForms10.EDIT.app.0.141b42a_r7_ad1", "txtVal", app_uia)
    cmd_handle = find_child_by_class(panel_handle, "WindowsForms10.BUTTON.app.0.141b42a_r7_ad1", "cmdAdd", app_uia)

    txt_val = app_win32.window(handle=txt_handle).wrapper_object()
    cmd_add = app_win32.window(handle=cmd_handle).wrapper_object()

    if not values:
        raise ValueError("Populate_Value_List requires at least one value")

    for value in (str(item) for item in values):
        txt_val.set_text(value)
        cmd_add.click()

    print(f"Completed interaction loop in {time.perf_counter() - start:6.3f} s")


def Clear_Value_List() -> None:
    start = time.perf_counter()

    # Connect to main app window via UIA
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    # Attach Win32 backend for same process
    app_win32 = Application(backend="win32").connect(process=app_uia.process)
    value_list_uia = main_uia.child_window(auto_id="lstValues", control_type="List")

    # Locate list control and wrap for Win32 ops
    value_list = app_win32.window(handle=value_list_uia.handle).wrapper_object()

    # Locate numeric panel container
    pnl_numeric = main_uia.child_window(auto_id="pnlNumeric", control_type="Pane")
    panel_handle = pnl_numeric.handle

    # Find and wrap Delete button inside panel
    cmdDelete_handle = find_child_by_class(panel_handle, "WindowsForms10.BUTTON.app.0.141b42a_r7_ad1", "cmdDelete", app_uia)
    cmdDelete = app_win32.window(handle=cmdDelete_handle).wrapper_object()


    list_length = len(value_list.item_texts())

    for i in range(list_length, 0, -1):
        value_list.click()
        cmdDelete.click()


    #print(f"{action} in {time.perf_counter() - start:6.3f} s")


def Example():
    start = time.perf_counter()

    # Connect to main app window via UIA
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    # Attach Win32 backend for same process
    app_win32 = Application(backend="win32").connect(process=app_uia.process)

    # Locate Pane button is in
    pnl_numeric = main_uia.child_window(auto_id="pnlNumeric", control_type="Pane")
    panel_handle = pnl_numeric.handle

    # Find and wrap button inside panel
    cmdDelete_handle = find_child_by_class(panel_handle, "WindowsForms10.BUTTON.app.0.141b42a_r7_ad1", "cmdDelete", app_uia)
    cmdDelete = app_win32.window(handle=cmdDelete_handle).wrapper_object()

    #Click Button
    cmdDelete.click()

    print(f"Example in {time.perf_counter() - start:6.3f} s")

def Open_Sensitivity_Analysis():
    start = time.perf_counter()

    # Connect to main app window via UIA
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    # Attach Win32 backend for same process
    app_win32 = Application(backend="win32").connect(process=app_uia.process)

    #main_uia.print_control_identifiers()

    # Locate Pane button is in
    pnl_menu = main_uia.child_window(title="Tools", control_type="MenuItem")
    Sensitivity_Analysis= pnl_menu.child_window(title="Sensitivity Analysis...", control_type="MenuItem")

    #Click Button
    pnl_menu.invoke()
    Sensitivity_Analysis.click_input()

    print(f"Open_Sensitivity_Analysis in {time.perf_counter() - start:6.3f} s")


def Open_Template():
    start = time.perf_counter()

    # Connect to main app window via UIA
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    #main_uia.print_control_identifiers()

    sensitivity_window = main_uia.child_window(auto_id="frmOrphSensitivity", control_type="Window")
    menu_bar = sensitivity_window.child_window(auto_id="MenuStrip1", control_type="MenuBar")
    file_menu = menu_bar.child_window(title="File", control_type="MenuItem").wrapper_object()
    file_menu.expand()

    deadline = time.perf_counter() + .10
    while time.perf_counter() < deadline:
        for child in file_menu.children():
            if child.element_info.name == "Open Template":
                child.click_input()
                print(f"Open_Template in {time.perf_counter() - start:6.3f} s")
                return
        #time.sleep(0.01)

    raise RuntimeError("Timeout locating 'Open Template' menu item")


def Set_Parameters_RIH_Old():
    start = time.perf_counter()

    # Connect to main app window via UIA
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    # Attach Win32 backend for same process
    app_win32 = Application(backend="win32").connect(process=app_uia.process)

    # Locate BHA depth button
    BHA_depth_numeric = main_uia.child_window(auto_id="chkDepth", title="BHA depth")
    BHA_depth_handle = BHA_depth_numeric.handle
    BHA_depth_Button = app_win32.window(handle=BHA_depth_handle).wrapper_object()

    # Locate Friction Factor Button
    friction_factor_numeric = main_uia.child_window(auto_id="chkFriction", title="Friction factor")
    friction_factor_handle = friction_factor_numeric.handle
    friction_factor_Button = app_win32.window(handle=friction_factor_handle).wrapper_object()

    # Locate Pipe Fluid Density Button
    Pipe_Fluid_Density = main_uia.child_window(auto_id="chkPipeFluidDens", title="Pipe fluid density")
    Pipe_Fluid_Density_handle = Pipe_Fluid_Density.handle
    Pipe_Fluid_Density_Button = app_win32.window(handle=Pipe_Fluid_Density_handle).wrapper_object()

    # Locate FOE RIH Button
    FOE_RIH = main_uia.child_window(auto_id="chkFOE", title="RIH")
    FOE_RIH_handle = FOE_RIH.handle
    FOE_RIH_Button = app_win32.window(handle=FOE_RIH_handle).wrapper_object()


    #Click Button
    BHA_depth_Button.click()
    friction_factor_Button.click()
    Pipe_Fluid_Density_Button.click()
    FOE_RIH_Button.click()
    Parameters_Minimum_Surface_Weight_During_RIH()


    print(f"Set_Parameters_RIH in {time.perf_counter() - start:6.3f} s")

def Set_Parameters_RIH():
    start = time.perf_counter()
    Parameters_Depth(checked=True)
    Parameters_FF(checked=False)
    Parameters_Pipe_fluid_density(checked=True)
    Parameters_RIH(checked=True)
    Parameters_Minimum_Surface_Weight_During_RIH(checked=True)

    print(f"Set_Parameters_RIH in {time.perf_counter() - start:6.3f} s")
