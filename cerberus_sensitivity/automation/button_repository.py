from __future__ import annotations

import re
from ctypes import windll, wintypes
from pathlib import Path
import time
from typing import Any, Mapping, Sequence
import warnings
import win32gui
import win32con

import pandas as pd
from pywinauto import Application, timings
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_element_info import UIAElementInfo

from .button_bridge_CSharp_to_py import collect_table, invoke_button, set_button_value

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

_DEFAULT_DUMP_PATH = (
    Path(__file__).resolve().parents[2] / "inspect_dumps" / "Windows_Inspect_Dump.txt"
)


# ---------------------------------------------------------------------------
# Data helpers
# ---------------------------------------------------------------------------

def _table_to_dataframe(
    table_payload: Mapping[str, Any],
    *,
    headers_override: Sequence[str] | None = None,
) -> pd.DataFrame:
    """Convert the C# helper payload into a Pandas DataFrame."""
    rows = list(table_payload.get("Rows") or [])
    headers = list(headers_override or table_payload.get("Headers") or [])

    if rows and headers and len(headers) != len(rows[0]):
        if len(headers) < len(rows[0]):
            headers.extend(f"Column {index}" for index in range(len(headers), len(rows[0])))
        else:
            headers = headers[: len(rows[0])]

    return pd.DataFrame(rows, columns=headers or None)


def _headers_from_dump(button_key: str, dump_path: Path | None = None) -> list[str]:
    """Read the static Inspect dump to recover friendly column headers."""
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
    else:
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
                        raw_headers.append(" ".join(pending))
                        pending = []
                    else:
                        raw_headers.append(core)
                else:
                    pending.append(stripped.strip('"'))
            if pending:
                raw_headers.append(" ".join(pending))

    return [" ".join(value.split()) for value in raw_headers if value.strip()]


def _coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Attempt to coerce string columns into numeric dtype where possible."""
    numeric_pattern = r"([-+]?\d*\.?\d+)"
    for column in df.columns:
        extracted = df[column].astype(str).str.extract(numeric_pattern)[0]
        if extracted.notna().any():
            numeric = pd.to_numeric(extracted, errors="coerce")
            if numeric.notna().any():
                df[column] = numeric
    return df


def _set_checkbox_state(button_key: str, checked: bool, *, timeout: float = 90.0):
    """Toggle a WinForms checkbox by writing to its ValuePattern."""
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


def _ensure_parameters_tab(timeout: float = 90.0) -> None:
    """Guarantee the Parameters tab is selected before acting on checkboxes."""
    Sensitivity_Setting_Parameters(timeout=timeout)


def _ensure_outputs_tab(timeout: float = 90.0) -> None:
    """Guarantee the Outputs tab is selected before acting on checkboxes."""
    Sensitivity_Setting_Outputs(timeout=timeout)


def _is_generic_header(name: object) -> bool:
    """Detect placeholder headers returned by the C# helper."""
    text = str(name or "").strip().lower()
    return not text or text == "#" or text.startswith("column")


def _normalize_uia_name(value: str | None) -> str:
    """Collapse whitespace/newlines and lowercase for reliable matching."""
    return " ".join(str(value or "").split()).lower()


# ---------------------------------------------------------------------------
# Navigation helpers and checkbox toggles
# ---------------------------------------------------------------------------

def button_Sensitivity_Analysis(timeout: float = 180.0):
    """Navigate to Tools > Sensitivity Analysis... (button1)."""
    return invoke_button("button1", timeout=timeout)


def Sensitivity_Setting_Parameters(timeout: float = 90.0):
    """Select the Parameters pane tab (button27)."""
    return invoke_button("button27", timeout=timeout)


def Sensitivity_Setting_Outputs(timeout: float = 90.0):
    """Select the Outputs pane tab (button7)."""
    return invoke_button("button7", timeout=timeout)


def Parameters_Pipe_fluid_density(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read Pipe Fluid Density (button5)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button5", timeout=timeout)
    return _set_checkbox_state("button5", checked, timeout=timeout)


def Parameters_Depth(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read BHA Depth (button32)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button32", timeout=timeout)
    return _set_checkbox_state("button32", checked, timeout=timeout)


def Parameters_FF(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read Friction Factor (button33)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button33", timeout=timeout)
    return _set_checkbox_state("button33", checked, timeout=timeout)


def Parameters_POOH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read the POOH parameter checkbox (button6)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button6", timeout=timeout)
    return _set_checkbox_state("button6", checked, timeout=timeout)


def Parameters_RIH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read the RIH parameter checkbox (button26)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button26", timeout=timeout)
    return _set_checkbox_state("button26", checked, timeout=timeout)


def Parameters_Maximum_Surface_Weight_During_POOH(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or read Max surface weight during POOH (button8)."""
    _ensure_outputs_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button8", timeout=timeout)
    return _set_checkbox_state("button8", checked, timeout=timeout)


def Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or read Max pipe stress during POOH (% of YS) (button9)."""
    _ensure_outputs_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button9", timeout=timeout)
    return _set_checkbox_state("button9", checked, timeout=timeout)


def Parameters_Minimum_Surface_Weight_During_RIH(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or read Min surface weight during RIH (button23)."""
    _ensure_outputs_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button23", timeout=timeout)
    return _set_checkbox_state("button23", checked, timeout=timeout)


def Setup_POOH(timeout: float = 90.0) -> None:
    """Apply the checkbox combination required for POOH runs."""
    Parameters_POOH(checked=True, timeout=timeout)
    Parameters_RIH(checked=False, timeout=timeout)
    Parameters_Maximum_Surface_Weight_During_POOH(checked=True, timeout=timeout)
    Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(checked=True, timeout=timeout)
    Parameters_Minimum_Surface_Weight_During_RIH(checked=False, timeout=timeout)


def Set_Parameters_RIH(timeout: float = 90.0) -> None:
    """Apply the checkbox combination required for RIH runs."""
    Parameters_Depth(checked=True, timeout=timeout)
    Parameters_FF(checked=False, timeout=timeout)
    Parameters_Pipe_fluid_density(checked=True, timeout=timeout)
    Parameters_RIH(checked=True, timeout=timeout)
    Parameters_Minimum_Surface_Weight_During_RIH(checked=True, timeout=timeout)


# ---------------------------------------------------------------------------
# Parameter Matrix interaction
# ---------------------------------------------------------------------------

def Parameter_Matrix_Wizard(timeout: float = 90.0):
    """Open the Parameter Matrix wizard (button10)."""
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrphSensitivity", control_type="Window", top_level_only=False)

    cmdMatrix_handle = main_uia.child_window(auto_id="cmdMatrix")
    cmdMatrix=cmdMatrix_handle.wrapper_object()
    return cmdMatrix.click_input()


def Parameter_Matrix_BHA_Depth_Row0(timeout: float = 60.0):
    app = Application(backend="uia").connect(auto_id="frmOrpheus")
    table = app.window(
        auto_id="frmSensitivityMatrix",
        control_type="Window",
        top_level_only=False,
    ).child_window(auto_id="grdVal", control_type="Table").wrapper_object()

    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 1)))
    cell.click_input()


def Parameter_Matrix_PFD_Row0(timeout: float = 60.0):
    app = Application(backend="uia").connect(auto_id="frmOrpheus")
    table = app.window(
        auto_id="frmSensitivityMatrix",
        control_type="Window",
        top_level_only=False,
    ).child_window(auto_id="grdVal", control_type="Table").wrapper_object()

    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 18)))
    cell.click_input()


def Parameter_Matrix_FOE_POOH_Row0(timeout: float = 60.0):
    app = Application(backend="uia").connect(auto_id="frmOrpheus")
    table = app.window(
        auto_id="frmSensitivityMatrix",
        control_type="Window",
        top_level_only=False,
    ).child_window(auto_id="grdVal", control_type="Table").wrapper_object()

    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 22)))
    cell.click_input()


def Parameter_Matrix_FOE_RIH_Row0(timeout: float = 60.0):
    app = Application(backend="uia").connect(auto_id="frmOrpheus")
    table = app.window(
        auto_id="frmSensitivityMatrix",
        control_type="Window",
        top_level_only=False,
    ).child_window(auto_id="grdVal", control_type="Table").wrapper_object()

    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 21)))
    cell.click_input()


# ---------------------------------------------------------------------------
# Value list editing helpers
# ---------------------------------------------------------------------------

def find_child_by_class(parent: int, class_name: str, automation_id: str, app_uia: Application) -> int:
    """Locate a child HWND by class name and automation id."""
    child = _FIND_WINDOW_EX(parent, 0, class_name, None)
    while child:
        if app_uia.window(handle=child).element_info.automation_id == automation_id:
            return child
        child = _FIND_WINDOW_EX(parent, child, class_name, None)
    raise RuntimeError(f"Unable to locate control '{automation_id}'.")


def Clear_Value_List() -> None:
    """Remove all values from the Parameter Matrix value list."""
    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    app_win32 = Application(backend="win32").connect(process=app_uia.process)
    value_list_uia = main_uia.child_window(auto_id="lstValues", control_type="List")
    value_list = app_win32.window(handle=value_list_uia.handle).wrapper_object()

    pnl_numeric = main_uia.child_window(auto_id="pnlNumeric", control_type="Pane")
    delete_handle = find_child_by_class(
        pnl_numeric.handle,
        "WindowsForms10.BUTTON.app.0.141b42a_r7_ad1",
        "cmdDelete",
        app_uia,
    )
    delete_button = app_win32.window(handle=delete_handle).wrapper_object()

    for _ in range(len(value_list.item_texts())):
        value_list.click()
        delete_button.click()


def Populate_Value_List(values: Sequence[str]) -> None:
    """Populate the Parameter Matrix value list with the provided values."""
    normalized_values = [str(item) for item in values]
    if not normalized_values:
        raise ValueError("Populate_Value_List requires at least one value.")

    app_uia = Application(backend="uia").connect(auto_id="frmOrpheus")
    main_uia = app_uia.window(auto_id="frmOrpheus", control_type="Window")

    app_win32 = Application(backend="win32").connect(process=app_uia.process)
    pnl_numeric = main_uia.child_window(auto_id="pnlNumeric", control_type="Pane")

    txt_handle = find_child_by_class(
        pnl_numeric.handle,
        "WindowsForms10.EDIT.app.0.141b42a_r7_ad1",
        "txtVal",
        app_uia,
    )
    add_handle = find_child_by_class(
        pnl_numeric.handle,
        "WindowsForms10.BUTTON.app.0.141b42a_r7_ad1",
        "cmdAdd",
        app_uia,
    )

    txt_val = app_win32.window(handle=txt_handle).wrapper_object()
    cmd_add = app_win32.window(handle=add_handle).wrapper_object()

    for value in normalized_values:
        txt_val.set_text(value)
        cmd_add.click()


def Edit_cmdOK(timeout: float = 60.0):
    """Confirm the value list edits (button16)."""
    return invoke_button("button16", timeout=timeout)


# ---------------------------------------------------------------------------
# Result collection helpers
# ---------------------------------------------------------------------------

def Sensitivity_Analysis_Calculate(timeout: float = 60.0):
    """Trigger the PAD Calculate button (button22)."""
    return invoke_button("button22", timeout=timeout)


def Sensitivity_Table(timeout: float = 90.0) -> pd.DataFrame:
    """Return the Sensitivity results grid (button24) as a DataFrame."""
    table_data = collect_table("button24", timeout=timeout)
    df = _table_to_dataframe(table_data)

    if df.columns.tolist() and df.columns[0] == "#":
        df = df.drop(columns="#")

    df = df.loc[:, df.apply(lambda col: col.astype(str).str.strip().ne("").any())]

    dump_headers = _headers_from_dump("button25") or _headers_from_dump("button24")
    if dump_headers and all(_is_generic_header(column) for column in df.columns):
        cleaned = [header for header in dump_headers if header and header != "#"]
        if cleaned:
            rename_map = dict(zip(df.columns, cleaned[: len(df.columns)]))
            df = df.rename(columns=rename_map)

    df = df.loc[:, df.apply(lambda col: col.astype(str).str.strip().ne("").any())]
    df = _coerce_numeric_columns(df)
    return df


def Sensitivity_Parameter_ok(timeout: float = 60.0):
    """Confirm the Parameter Matrix edits (button29)."""
    return invoke_button("button29", timeout=timeout)

def Example() -> None:
    """
    Connects to the application and automates entering numbers.
    """
    start_time = time.time()

    app = Application(backend="uia").connect(auto_id="frmOrpheus", timeout=20)
    print(f"Connected in {time.time() - start_time:.2f}s")

    # --- THE CORRECTED LOGIC ---

    # 1. Create the RECIPE for the parent form. DO NOT call .wrapper_object() yet.
    # This is just a plan for how to find it.
    sensitivity_form_spec = app.window(auto_id="frmOrphSensitivity",top_level_only=False)

    # 2. Now, use the parent recipe to create the full recipes for the child controls.
    txt_val_spec = sensitivity_form_spec.child_window(auto_id="txtVal")
    cmd_add_spec = sensitivity_form_spec.child_window(auto_id="cmdAdd")

    # 3. NOW we execute the recipes to get the final, cached controls.
    # The slow search happens here, just once for each control.
    print("Finding and caching controls...")
    txt_val = txt_val_spec.wrapper_object()
    cmd_add = cmd_add_spec.wrapper_object()
    print(f"Controls found and cached in {time.time() - start_time:.2f}s")

    # 4. The loop is fast because it uses the cached "final dish" objects.
    for i in range(1, 11):
        txt_val.set_text(str(i))
        cmd_add.click()

    print(f"\nSuccessfully entered numbers 1 through 10.")
    print(f"Total execution time: {time.time() - start_time:.2f}s")

