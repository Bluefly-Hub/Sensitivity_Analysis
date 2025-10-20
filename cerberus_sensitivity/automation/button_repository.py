from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Any, Mapping, Sequence

import pandas as pd

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


def Parameters_POOH(checked: bool | None = None, timeout: float = 90.0):
    """Toggle or set the POOH parameter checkbox (button6)."""
    _ensure_parameters_tab(timeout=timeout)
    if checked is None:
        return invoke_button("button6", timeout=timeout)
    return _set_checkbox_state("button6", checked, timeout=timeout)


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
    return invoke_button("button17", timeout=timeout)


def Parameter_Matrix_FOE_POOH_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for FOE row 0 (button19)."""
    return invoke_button("button19", timeout=timeout)

def Parameter_Matrix_FOE_RIH_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for FOE row 0 (button28)."""
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

