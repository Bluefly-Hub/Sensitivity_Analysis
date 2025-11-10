from __future__ import annotations

import re
from ctypes import windll, wintypes
from pathlib import Path
from typing import Any, Mapping, Sequence
import warnings
import comtypes.client
from comtypes.gen.UIAutomationClient import IUIAutomation, TreeScope_Descendants

import pandas as pd
from pywinauto import Application, timings
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_element_info import UIAElementInfo

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
# Cached Application Connection
# ---------------------------------------------------------------------------

class _AppConnection:
    """Singleton class to cache the application connection and root element."""
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
        
        # Connect once and reuse the app connection
        self.app = Application(backend="uia").connect(auto_id="frmOrpheus")
        # Get root element for fast searches
        self.root = self.app.top_window().element_info.element
        self._initialized = True
    
    def get_root(self):
        """Get the cached root element."""
        return self.root
    
    def get_app(self):
        """Get the cached application."""
        return self.app
    
    def refresh(self):
        """Refresh the connection if needed."""
        try:
            # Test if connection is still valid
            _ = self.app.top_window()
        except Exception:
            # Reconnect
            self.app = Application(backend="uia").connect(auto_id="frmOrpheus")
            self.root = self.app.top_window().element_info.element


def _get_app_root():
    """Get the cached root element for fast searches."""
    return _AppConnection().get_root()


def _get_app():
    """Get the cached application connection."""
    return _AppConnection().get_app()


# ---------------------------------------------------------------------------
# Fast element search helpers
# ---------------------------------------------------------------------------

def find_element_fast(root_element, automation_id, found_index=0):
    """
    Fast element search using direct UIA API
    10x faster than pywinauto's window() search
    """
    uia = comtypes.client.GetModule('UIAutomationCore.dll')
    iuia = comtypes.client.CreateObject('{ff48dba4-60ef-4201-aa87-54103eef594e}', interface=IUIAutomation)
    
    condition = iuia.CreatePropertyCondition(30011, automation_id)  # AutomationId
    
    if found_index == 0:
        # Just find first
        element = root_element.FindFirst(TreeScope_Descendants, condition)
        return UIAWrapper(UIAElementInfo(element)) if element else None
    else:
        # Find all and return specific index
        elements_array = root_element.FindAll(TreeScope_Descendants, condition)
        if found_index < elements_array.Length:
            element = elements_array.GetElement(found_index)
            return UIAWrapper(UIAElementInfo(element))
        return None

def find_element_by_title(root_element, title):
    """Fast search by title/name"""
    uia = comtypes.client.GetModule('UIAutomationCore.dll')
    iuia = comtypes.client.CreateObject('{ff48dba4-60ef-4201-aa87-54103eef594e}', interface=IUIAutomation)
    
    condition = iuia.CreatePropertyCondition(30005, title)  # Name property
    element = root_element.FindFirst(TreeScope_Descendants, condition)
    return UIAWrapper(UIAElementInfo(element)) if element else None

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

def _get_checkbox_and_toggle(automation_id: str, checked: bool | None):
    """Helper to get checkbox and toggle its state if needed."""
    root = _get_app_root()
    checkbox = find_element_fast(root, automation_id)
    
    if checked is None:
        # Read mode - just click to toggle
        #checkbox.set_focus()
        checkbox.toggle()
    else:
        # Set mode - check current state and only click if different
        try:
            current_state = checkbox.iface_toggle.CurrentToggleState
            is_checked = (current_state == 1)  # 1 = On, 0 = Off
            
            if is_checked != checked:
                #checkbox.set_focus()
                checkbox.toggle()
        except Exception:
            # Fallback if toggle pattern not available
            #checkbox.set_focus()
            checkbox.toggle()


def button_Sensitivity_Analysis():
    """Navigate to Tools > Sensitivity Analysis... (button1)."""
    root = _get_app_root()
    MenuStrip1 = find_element_fast(root, "MenuStrip1").element_info.element
    tools_menu = find_element_by_title(MenuStrip1, "Tools")
    tools_menu.set_focus()
    tools_menu.click_input()
    
    # Now find and click "Sensitivity Analysis..." menu item
    sensitivity_item = find_element_by_title(MenuStrip1, "Sensitivity Analysis...")
    sensitivity_item.set_focus()
    sensitivity_item.click_input()


def Sensitivity_Setting_Parameters(timeout: float = 90.0):
    """Select the Parameters pane tab (button27)."""
    root = _get_app_root()
    params_tab = find_element_by_title(root, "Parameters")
    
    # Check if already selected - tabs have IsSelected property
    try:
        is_selected = params_tab.get_selection_item_pattern().CurrentIsSelected
        if not is_selected:
            params_tab.select()
    except Exception:
        # If we can't check, just try selecting (it won't error if already selected)
        try:
            params_tab.select()
        except Exception:
            # Already selected, ignore the error
            pass


def Sensitivity_Setting_Outputs(timeout: float = 90.0):
    """Select the Outputs pane tab (button7)."""
    root = _get_app_root()
    outputs_tab = find_element_by_title(root, "Outputs")
    
    # Check if already selected - tabs have IsSelected property
    try:
        is_selected = outputs_tab.get_selection_item_pattern().CurrentIsSelected
        if not is_selected:
            outputs_tab.select()
    except Exception:
        # If we can't check, just try selecting (it won't error if already selected)
        try:
            outputs_tab.select()
        except Exception:
            # Already selected, ignore the error
            pass


def Parameters_Pipe_fluid_density(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read Pipe Fluid Density (button5)."""
    #_ensure_parameters_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkPipeFluidDens", checked)



def Parameters_Depth(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read BHA Depth (button32)."""
    #_ensure_parameters_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkDepth", checked)


def Parameters_FF(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read Friction Factor (button33)."""
    #_ensure_parameters_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkFriction", checked)


def Parameters_POOH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read the POOH parameter checkbox (button6)."""
    #_ensure_parameters_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkFOE_POOH", checked)


def Parameters_RIH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or read the RIH parameter checkbox (button26)."""
    #_ensure_parameters_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkFOE", checked)


def Parameters_Maximum_Surface_Weight_During_POOH(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or read Max surface weight during POOH (button8)."""
    #_ensure_outputs_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkPOOH_MaxSW", checked)


def Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or read Max pipe stress during POOH (% of YS) (button9)."""
    #_ensure_outputs_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkPOOH_MaxYield", checked)


def Parameters_Minimum_Surface_Weight_During_RIH(
    checked: bool | None = None, timeout: float = 90.0
):
    """Toggle or read Min surface weight during RIH (button23)."""
    #_ensure_outputs_tab(timeout=timeout)
    _get_checkbox_and_toggle("chkRIH_MinSW", checked)


def Setup_POOH(timeout: float = 90.0) -> None:
    """Apply the checkbox combination required for POOH runs."""
    _ensure_parameters_tab(timeout=timeout)
    Parameters_POOH(checked=True, timeout=timeout)
    Parameters_RIH(checked=False, timeout=timeout)
    Parameters_Depth(checked=True, timeout=timeout)
    Parameters_FF(checked=False, timeout=timeout)
    Parameters_Pipe_fluid_density(checked=True, timeout=timeout)
    _ensure_outputs_tab(timeout=timeout)
    Parameters_Maximum_Surface_Weight_During_POOH(checked=True, timeout=timeout)
    Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(checked=True, timeout=timeout)
    Parameters_Minimum_Surface_Weight_During_RIH(checked=False, timeout=timeout)


def Set_Parameters_RIH(timeout: float = 90.0) -> None:
    """Apply the checkbox combination required for RIH runs."""
    _ensure_parameters_tab(timeout=timeout)
    Parameters_POOH(checked=False, timeout=timeout)
    Parameters_RIH(checked=True, timeout=timeout)
    Parameters_Depth(checked=True, timeout=timeout)
    Parameters_FF(checked=False, timeout=timeout)
    Parameters_Pipe_fluid_density(checked=True, timeout=timeout)
    _ensure_outputs_tab(timeout=timeout)
    Parameters_Maximum_Surface_Weight_During_POOH(checked=False, timeout=timeout)
    Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(checked=False, timeout=timeout)
    Parameters_Minimum_Surface_Weight_During_RIH(checked=True, timeout=timeout)






# ---------------------------------------------------------------------------
# Parameter Matrix interaction
# ---------------------------------------------------------------------------

def Parameter_Matrix_Wizard(timeout: float = 90.0):
    """Open the Parameter Matrix wizard (button10)."""
    root = _get_app_root()
    main_uia = find_element_fast(root, "frmOrphSensitivity")
    find_element_fast(main_uia.element_info.element, "cmdMatrix").click_input()


def Parameter_Matrix_BHA_Depth_Row0(timeout: float = 60.0):
    """Select and activate BHA Depth cell in Parameter Matrix (Row 0, Column 1)."""
    root = _get_app_root()
    matrix_window = find_element_fast(root, "frmSensitivityMatrix")
    table = find_element_fast(matrix_window.element_info.element, "grdVal")
    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 1)))
    cell.click_input()  # Double-click to open value editor


def Parameter_Matrix_PFD_Row0(timeout: float = 60.0):
    """Select and activate Pipe Fluid Density cell in Parameter Matrix (Row 0, Column 18)."""
    root = _get_app_root()
    matrix_window = find_element_fast(root, "frmSensitivityMatrix")
    table = find_element_fast(matrix_window.element_info.element, "grdVal")
    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 18)))
    cell.click_input()


def Parameter_Matrix_FOE_POOH_Row0(timeout: float = 60.0):
    """Select and activate FOE POOH cell in Parameter Matrix (Row 0, Column 22)."""
    root = _get_app_root()
    matrix_window = find_element_fast(root, "frmSensitivityMatrix")
    table = find_element_fast(matrix_window.element_info.element, "grdVal")
    cell = UIAWrapper(UIAElementInfo(table.iface_grid.GetItem(0, 22)))
    cell.click_input()


def Parameter_Matrix_FOE_RIH_Row0(timeout: float = 60.0):
    """Select and activate FOE RIH cell in Parameter Matrix (Row 0, Column 21)."""
    root = _get_app_root()
    matrix_window = find_element_fast(root, "frmSensitivityMatrix")
    table = find_element_fast(matrix_window.element_info.element, "grdVal")
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
    app_uia = _get_app()
    root = _get_app_root()

    # Use fast lookup for list and delete button
    value_list_uia = find_element_fast(root, "lstValues")
    delete_button_uia = find_element_fast(root, "cmdDelete")
    
    # Convert to Win32 for item_texts() method
    app_win32 = Application(backend="win32").connect(process=app_uia.process)
    value_list = app_win32.window(handle=value_list_uia.handle).wrapper_object()
    delete_button = app_win32.window(handle=delete_button_uia.handle).wrapper_object()

    for _ in range(len(value_list.item_texts())):
        value_list.click()
        delete_button.click()


def Populate_Value_List(values: Sequence[str]) -> None:
    """Populate the Parameter Matrix value list with the provided values."""
    normalized_values = [str(item) for item in values]
    if not normalized_values:
        raise ValueError("Populate_Value_List requires at least one value.")

    app_uia = _get_app()
    root = _get_app_root()

    # Use fast lookup for text box and add button
    txt_val_uia = find_element_fast(root, "txtVal")
    cmd_add_uia = find_element_fast(root, "cmdAdd")
    
    # Convert to Win32 for set_text() method
    app_win32 = Application(backend="win32").connect(process=app_uia.process)
    txt_val = app_win32.window(handle=txt_val_uia.handle).wrapper_object()
    cmd_add = app_win32.window(handle=cmd_add_uia.handle).wrapper_object()

    for value in normalized_values:
        txt_val.set_text(value)
        cmd_add.click()


def Edit_cmdOK(timeout: float = 60.0):
    """Confirm the value list edits (button16)."""
    root = _get_app_root()
    find_element_fast(root, "cmdOK").click_input()
    #return invoke_button("button16", timeout=timeout)


# ---------------------------------------------------------------------------
# Result collection helpers
# ---------------------------------------------------------------------------

def Sensitivity_Analysis_Calculate(timeout: float = 60.0):
    """Trigger the PAD Calculate button (button22)."""
    root = _get_app_root()
    element = find_element_fast(root, "cmdCalc")
    element.iface_invoke.Invoke()


def Sensitivity_Table(timeout: float = 90.0) -> pd.DataFrame:
    """Extract Sensitivity grid data and return as pandas DataFrame."""
    # Get the app and refresh root to ensure we have current window
    app = _get_app()
    root = app.top_window().element_info.element
    
    # Find the sensitivity window and grid element
    sensitivity_window = find_element_fast(root, "frmOrphSensitivity")
    if not sensitivity_window:
        raise RuntimeError("Sensitivity Analysis window not found")
    
    sensitivity_element = sensitivity_window.element_info.element
    grid = find_element_fast(sensitivity_element, "grdSensitivityData")
    if not grid:
        raise RuntimeError("Sensitivity results grid 'grdSensitivityData' not found")
    
    # Access grid's children (rows)
    rows = grid.children()
    data = []
    headers = []
    
    for i, row in enumerate(rows):
        row_data = []
        cells = row.children()  # Get cells in the row
        for cell in cells:
            # Try to get actual value instead of title
            try:
                # Try Value pattern first
                if hasattr(cell, 'iface_value') and cell.iface_value:
                    cell_text = cell.get_value()
                else:
                    # Fall back to legacy value
                    cell_text = cell.legacy_properties().get('Value', cell.window_text())
            except:
                # Last resort - use window text
                cell_text = cell.window_text()
            row_data.append(cell_text)
        
        # Check if this row looks like a header (first row or contains header-like content)
        is_header = (i == 0) or any(
            'BHA Depth' in str(cell) or 
            'Pipe Fluid Density' in str(cell) or 
            'Force on End' in str(cell) or
            '\r(' in str(cell) 
            for cell in row_data
        )
        
        if is_header:
            headers = row_data  # This is a header row
            # Don't add it to data
        else:
            # Only add non-header rows to data
            if headers:  # Only add data if we've found headers
                data.append(row_data)
    
    # If no headers detected, use first row as headers
    if not headers and data:
        headers = data[0]
        data = data[1:]
    
    # Create DataFrame
    df = pd.DataFrame(data, columns=headers if headers else None)
    
    # Drop first column if it's just row numbers (typically "#")
    if df.columns.tolist() and (df.columns[0] == "#" or df.columns[0] == ""):
        df = df.iloc[:, 1:]
    
    # Remove empty columns
    df = df.loc[:, df.apply(lambda col: col.astype(str).str.strip().ne("").any())]
    
    # Convert numeric columns
    df = _coerce_numeric_columns(df)
    
    return df


def Sensitivity_Parameter_ok(timeout: float = 60.0):
    """Confirm the Parameter Matrix edits (button29)."""
    root = _get_app_root()
    element = find_element_fast(root, "cmdOK")
    element.iface_invoke.Invoke()

