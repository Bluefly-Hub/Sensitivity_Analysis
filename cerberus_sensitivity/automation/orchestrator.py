from __future__ import annotations

import subprocess
import threading
import time
from dataclasses import dataclass
from typing import Any, Callable, Dict, Iterable, List, Sequence, Tuple

import pandas as pd

from .button_repository import (
    Edit_cmdAdd,
    Edit_cmdDelete,
    Edit_cmdOK,
    File_OpenTemplate,
    File_OpenTemplate_auto,
    Parameter_Matrix_BHA_Depth_Row0,
    Parameter_Matrix_FOE_POOH_Row0,
    Parameter_Matrix_FOE_RIH_Row0,
    Parameter_Matrix_PFD_Row0,
    Parameter_Matrix_Wizard,
    Parameter_Value_Editor_Set_Value,
    Parameters_Maximum_Surface_Weight_During_POOH,
    Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS,
    Parameters_Minimum_Surface_Weight_During_RIH,
    Parameters_POOH,
    Parameters_Pipe_fluid_density,
    Parameters_RIH,
    Sensitivity_Analysis_Calculate,
    Sensitivity_Setting_Outputs,
    Sensitivity_Table,
    Value_List_Item0,
    button_Sensitivity_Analysis,
    button_exit_wizard,
    Sensitivity_Parameter_ok,
)
from .inputs import SensitivityInputRow, SensitivityInputs
from .progress import ProgressReporter
from .run_plan import RunPlan, chunk_depths


MAX_ITERATIONS_PER_RUN = 200
RIH_MODE = "RIH"
POOH_MODE = "POOH"

_TEMPLATE_SELECTORS: Dict[str, Callable[..., object]] = {
    "auto": File_OpenTemplate_auto,
}

_PARAMETER_ROW_SELECTORS: Dict[str, Callable[..., object]] = {
    "BHA Depth": Parameter_Matrix_BHA_Depth_Row0,
    "Pipe Fluid Density": Parameter_Matrix_PFD_Row0,
    "Force on End - RIH": Parameter_Matrix_FOE_RIH_Row0,
    "Force on End - POOH": Parameter_Matrix_FOE_POOH_Row0,
}


@dataclass
class ScanResult:
    index: int
    payload: Dict[str, Any]


class CerberusOrchestrator:
    def __init__(self) -> None:
        pass

    def run_scan(
        self,
        progress: ProgressReporter,
        data_list: Sequence[Dict[str, Any] | SensitivityInputRow],
        start_index: int,
        cancel_event: threading.Event,
        template_name: str = "auto",
    ) -> tuple[List[Dict[str, Any]], Dict[str, pd.DataFrame]]:
        rows = [self._coerce_row(item) for item in data_list]
        inputs = SensitivityInputs.from_rows(rows)
        run_plan = self._build_run_plan(inputs)
        if not run_plan:
            raise ValueError(
                "No valid Cerberus sensitivity combinations were detected. Check that density, depth, and FOE columns contain numeric values."
            )
        total_rows = run_plan[-1].end_offset

        if start_index >= total_rows:
            progress.init(total_rows, template=template_name)
            empty_outputs = {RIH_MODE: pd.DataFrame(), POOH_MODE: pd.DataFrame()}
            progress.done([row.__dict__ for row in rows], empty_outputs)
            return [row.__dict__ for row in rows], empty_outputs

        rih_frames: List[pd.DataFrame] = []
        pooh_frames: List[pd.DataFrame] = []
        parameter_value_cache: Dict[str, Tuple[float, ...]] = {}

        progress.init(total_rows, template=template_name)
        _open_sensitivity_analysis()
        _load_template(template_name)

        current_mode: str | None = None

        for plan in run_plan:
            if cancel_event.is_set():
                break
            if start_index >= plan.end_offset:
                continue

            if plan.mode != current_mode:
                _configure_parameters_for_mode(plan.mode)
                _configure_outputs_for_mode(plan.mode)
                current_mode = plan.mode

            _open_parameter_matrix()
            # BHA depth must be loaded first so subsequent parameters reference the correct slice
            _update_parameter_values("BHA Depth", plan.depth_values, cache=parameter_value_cache)
            if plan.include_pipe_density:
                _update_parameter_values(
                    "Pipe Fluid Density", inputs.pipe_fluid_densities, cache=parameter_value_cache
                )
            if plan.include_force_on_end:
                foe_values = inputs.stretch_foe_rih if plan.mode == "RIH" else inputs.stretch_foe_pooh
                caption = "Force on End - RIH" if plan.mode == "RIH" else "Force on End - POOH"
                # FOE lists differ by mode, so always refresh instead of using cache
                _update_parameter_values(caption, foe_values)
            _close_parameter_matrix()

            table_df = _recalc_and_collect_table()
            table_rows = table_df.to_dict(orient="records")
            new_rows_for_mode: List[Dict[str, Any]] = []
            for idx, row in enumerate(table_rows):
                global_index = plan.offset + idx
                if global_index < start_index:
                    continue
                if cancel_event.is_set():
                    break
                row_payload = {"mode": plan.mode, **row}
                progress.row(global_index, row_payload)
                stored_row = {"index": global_index, **row}
                new_rows_for_mode.append(stored_row)
            if cancel_event.is_set():
                break
            if new_rows_for_mode:
                frame = pd.DataFrame(new_rows_for_mode)
                if plan.mode == RIH_MODE:
                    rih_frames.append(frame)
                else:
                    pooh_frames.append(frame)

        rih_df = pd.concat(rih_frames, ignore_index=True) if rih_frames else pd.DataFrame()
        pooh_df = pd.concat(pooh_frames, ignore_index=True) if pooh_frames else pd.DataFrame()
        output_frames = {RIH_MODE: rih_df, POOH_MODE: pooh_df}
        progress.done([row.__dict__ for row in rows], output_frames)
        return [row.__dict__ for row in rows], output_frames

    def _coerce_row(self, item: Dict[str, Any] | SensitivityInputRow) -> SensitivityInputRow:
        if isinstance(item, SensitivityInputRow):
            return item
        return SensitivityInputRow(
            pipe_fluid_density=self._maybe_float(item.get("pipe_fluid_density")),
            sleeve_number=item.get("sleeve_number"),
            depth=self._maybe_float(item.get("depth")),
            stretch_foe_rih=self._maybe_float(item.get("stretch_foe_rih")),
            stretch_foe_pooh=self._maybe_float(item.get("stretch_foe_pooh")),
        )

    def _maybe_float(self, value: Any) -> float | None:
        if value in (None, "", "-"):
            return None
        if isinstance(value, str):
            cleaned = value.strip().replace(",", "")
            if cleaned in ("", "-"):
                return None
            value = cleaned
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    def _build_run_plan(self, inputs: SensitivityInputs) -> List[RunPlan]:
        plan: List[RunPlan] = []
        offset = 0
        if inputs.pipe_fluid_densities and inputs.stretch_foe_rih and inputs.depths:
            plan, offset = self._extend_plan(plan, offset, inputs, mode="RIH")
        if inputs.pipe_fluid_densities and inputs.stretch_foe_pooh and inputs.depths:
            plan, offset = self._extend_plan(plan, offset, inputs, mode="POOH")
        return plan

    def _extend_plan(
        self,
        plan: List[RunPlan],
        offset: int,
        inputs: SensitivityInputs,
        mode: str,
    ) -> tuple[List[RunPlan], int]:
        foe_values = inputs.stretch_foe_rih if mode == "RIH" else inputs.stretch_foe_pooh
        chunk_size = self._compute_depth_chunk(len(inputs.depths), len(inputs.pipe_fluid_densities), len(foe_values))
        depth_chunks = chunk_depths(inputs.depths, chunk_size) if inputs.depths else []
        for index, chunk in enumerate(depth_chunks):
            combo_count = len(inputs.pipe_fluid_densities) * len(foe_values) * len(chunk)
            plan.append(
                RunPlan(
                    mode=mode,
                    depth_values=chunk,
                    include_pipe_density=(index == 0),
                    include_force_on_end=(index == 0),
                    combo_count=combo_count,
                    offset=offset,
                )
            )
            offset += combo_count
        return plan, offset

    def _compute_depth_chunk(self, depth_count: int, density_count: int, foe_count: int) -> int:
        if depth_count == 0:
            return 0
        combos_per_depth = density_count * foe_count
        if combos_per_depth == 0:
            return depth_count
        chunk = depth_count
        while combos_per_depth * chunk > MAX_ITERATIONS_PER_RUN and chunk > 1:
            chunk -= 1
        return max(1, chunk)


def _open_sensitivity_analysis() -> None:
    button_Sensitivity_Analysis()
    #time.sleep(0.3)


def _load_template(template_name: str, timeout: float = 10.0) -> None:
    key = template_name.strip().lower()
    selector = _TEMPLATE_SELECTORS.get(key)
    if selector is None:
        raise ValueError(f"Unsupported template '{template_name}'.")

    File_OpenTemplate(timeout=timeout)
    #time.sleep(0.2)
    selector(timeout=timeout)
    #time.sleep(0.5)


def _configure_parameters_for_mode(mode: str) -> None:
    normalized = mode.strip().upper()
    if normalized not in {RIH_MODE, POOH_MODE}:
        raise ValueError(f"Unsupported mode '{mode}'.")

    Parameters_Pipe_fluid_density(checked=True)
    if normalized == RIH_MODE:
        Parameters_RIH(checked=True)
        Parameters_POOH(checked=False)
    else:
        Parameters_RIH(checked=False)
        Parameters_POOH(checked=True)
    ##time.sleep(0.2)


def _configure_outputs_for_mode(mode: str) -> None:
    normalized = mode.strip().upper()
    if normalized not in {RIH_MODE, POOH_MODE}:
        raise ValueError(f"Unsupported mode '{mode}'.")

    Sensitivity_Setting_Outputs()
    #time.sleep(0.1)

    if normalized == RIH_MODE:
        Parameters_Minimum_Surface_Weight_During_RIH(checked=True)
        Parameters_RIH(checked=True)
        Parameters_Pipe_fluid_density(checked=True)
        Parameters_Maximum_Surface_Weight_During_POOH(checked=False)
        Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(checked=False)
    else:
        Parameters_Minimum_Surface_Weight_During_RIH(checked=False)
        Parameters_Maximum_Surface_Weight_During_POOH(checked=True)
        Parameters_Maximum_pipe_stress_during_POOH_percent_of_YS(checked=True)
    #time.sleep(0.2)


def _open_parameter_matrix() -> None:
    Parameter_Matrix_Wizard()
    #time.sleep(0.5)


def _close_parameter_matrix() -> None:
    Sensitivity_Parameter_ok()
    #time.sleep(0.3)


def _update_parameter_values(
    parameter_caption: str,
    values: Iterable[float],
    ensure_clear: bool = True,
    cache: Dict[str, Tuple[float, ...]] | None = None,
) -> None:
    normalized_caption = parameter_caption.strip()
    normalized_key = normalized_caption

    sequence = [value for value in values if value is not None]
    if not sequence:
        if cache is not None:
            cache.pop(normalized_key, None)
        return

    sequence_tuple: Tuple[float, ...] = tuple(sequence)
    if cache is not None and cache.get(normalized_key) == sequence_tuple:
        return

    selector = _resolve_parameter_selector(parameter_caption)
    selector()
    #time.sleep(0.2)

    if ensure_clear:
        _clear_value_list()

    for value in sequence:
        Parameter_Value_Editor_Set_Value(str(value))
        #time.sleep(0.05)
        Edit_cmdAdd()
        #time.sleep(0.05)

    Edit_cmdOK()
    #time.sleep(0.2)

    if cache is not None:
        cache[normalized_key] = sequence_tuple


def _resolve_parameter_selector(caption: str):
    normalized = caption.strip()
    selector = _PARAMETER_ROW_SELECTORS.get(normalized)
    if selector is None:
        raise ValueError(f"No parameter matrix selector defined for '{caption}'.")
    return selector


def _clear_value_list(max_attempts: int = 1) -> None:
    attempts = 0
    while attempts < max_attempts:
        attempts += 1
        try:
            Value_List_Item0()
            #time.sleep(0.05)
        except subprocess.CalledProcessError:
            break
        try:
            Edit_cmdDelete()
            #time.sleep(0.05)
        except subprocess.CalledProcessError:
            break


def _recalc_and_collect_table(timeout: float = 120.0) -> pd.DataFrame:
    Sensitivity_Analysis_Calculate(timeout=timeout)
    ##time.sleep(1.0)
    table = Sensitivity_Table(timeout=timeout)
    if table is None or table.empty:
        raise RuntimeError("Sensitivity table did not return any data.")
    return table



