import re
from pathlib import Path
from typing import Any, Mapping, Sequence

import pandas as pd

from cerberus_sensitivity.automation.button_bridge_CSharp_to_py import (
    collect_table,
    invoke_button,
    list_buttons,
    set_button_value,
)

_DEFAULT_DUMP_PATH = Path(__file__).resolve().parent / "inspect_dumps" / "Windows_Inspect_Dump.txt"


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
    if not header_block_match:
        return []

    raw_headers = re.findall(r'"([^"]+)"', header_block_match.group(1))
    return [" ".join(value.split()) for value in raw_headers if value.strip()]

# def Sensitivity_Analysis():
#     """Invoke the Sensitivity Analysis menu via the compiled C# helper."""
#     ...
# Keeping the layout similar to your original script, but pruning unused imports.


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

# def Value_List_Depth_1(timeout: float = 60.0):
#     """Open the value list entry for Depth value 1 (button12)."""
#     return invoke_button("button12", timeout=timeout)

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


def Parameter_Matrix_FOE_Row0(timeout: float = 60.0):
    """Select the Parameter Matrix grid cell for FOE row 0 (button19)."""
    return invoke_button("button19", timeout=timeout)


def Value_List_Item0(timeout: float = 60.0):
    """Open the value list entry for FOE value 2 (button21)."""
    return invoke_button("button21", timeout=timeout)

def Sensitivity_Analysis_Calculate(timeout: float = 60.0):
    """Trigger the PAD Calculate button (button22)."""
    return invoke_button("button22", timeout=timeout)

def Parameters_Minimum_Surface_Weight_During_RIH(timeout: float = 90.0):
    """Toggle Minimum surface weight during RIH (button23)."""
    return invoke_button("button23", timeout=timeout)

def Sensitivity_Table(timeout: float = 90.0):
    """Sensitivity Table (button24)."""
    table_data = collect_table("button24", timeout=timeout)
    df = _table_to_dataframe(table_data)
    df = df.loc[:, df.apply(lambda col: col.astype(str).str.strip().ne("").any())]

    dump_headers = _headers_from_dump("button24")
    target_sequence = [
        "BHA Depth (ft)",
        "Pipe Fluid Density (lb/gal)",
        "Force on End - RIH (lbf)",
        "Lockup Depth (ft)",
        "Min Surface Wt - RIH (lbf)",
        "Max Surface Wt - POOH (lbf)",
        "Max Pipe Stress - POOH (% of YS)",
    ]

    if dump_headers and len(df.columns) == len(target_sequence) + 1:
        resolved_headers: list[str] = []
        for target in target_sequence:
            match = next((name for name in dump_headers if name.startswith(target.split(" (")[0])), target)
            resolved_headers.append(match)
        df.columns = ["#"] + resolved_headers
    elif dump_headers:
        df.columns = dump_headers[: len(df.columns)]

    if "#" in df.columns:
        df = df.drop(columns="#")
    df = _coerce_numeric_columns(df)
    print(df)
    return df

if __name__ == "__main__":
    # button_Sensitivity_Analysis()
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
    # Value_List_Item0()
    # Sensitivity_Analysis_Calculate()
    # Parameters_Minimum_Surface_Weight_During_RIH()
    Sensitivity_Table()


# list_buttons
