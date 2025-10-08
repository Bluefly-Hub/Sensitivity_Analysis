from __future__ import annotations

import threading
from dataclasses import dataclass
from typing import Any, Dict, List, Sequence

from .clipboard import parse_tabulated_tsv
from .inputs import SensitivityInputRow, SensitivityInputs
from .progress import ProgressReporter
from .run_plan import RunPlan, chunk_depths
from .ui_trigger import UITriggerAgent


MAX_ITERATIONS_PER_RUN = 200


@dataclass
class ScanResult:
    index: int
    payload: Dict[str, Any]


class CerberusOrchestrator:
    def __init__(self, trigger: UITriggerAgent) -> None:
        self.trigger = trigger

    def run_scan(
        self,
        progress: ProgressReporter,
        data_list: Sequence[Dict[str, Any] | SensitivityInputRow],
        start_index: int,
        cancel_event: threading.Event,
        template_name: str = "auto",
    ) -> tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        rows = [self._coerce_row(item) for item in data_list]
        inputs = SensitivityInputs.from_rows(rows)
        run_plan = self._build_run_plan(inputs)
        if not run_plan:
            raise ValueError("No valid Cerberus sensitivity combinations were detected. Check that density, depth, and FOE columns contain numeric values.")
        total_rows = run_plan[-1].end_offset

        if start_index >= total_rows:
            progress.init(total_rows, template=template_name)
            progress.done([row.__dict__ for row in rows], [])
            return [row.__dict__ for row in rows], []

        progress.init(total_rows, template=template_name)
        self.trigger.open_sensitivity_analysis()
        self.trigger.load_template(template_name)

        aggregated_results: List[Dict[str, Any]] = []

        for plan in run_plan:
            if cancel_event.is_set():
                break
            if start_index >= plan.end_offset:
                continue

            self.trigger.configure_outputs_for_mode(plan.mode)
            self.trigger.configure_for_mode(plan.mode)

            self.trigger.open_parameter_matrix()
            if plan.include_pipe_density:
                self.trigger.update_parameter_values("Pipe Fluid Density", inputs.pipe_fluid_densities)
            if plan.include_force_on_end:
                foe_values = inputs.stretch_foe_rih if plan.mode == "RIH" else inputs.stretch_foe_pooh
                caption = "Force on End - RIH" if plan.mode == "RIH" else "Force on End - POOH"
                self.trigger.update_parameter_values(caption, foe_values)
            self.trigger.update_parameter_values("BHA Depth", plan.depth_values)
            self.trigger.close_parameter_matrix()

            raw_text = self.trigger.recalc_and_copy_results()
            parsed_rows = parse_tabulated_tsv(raw_text)

            for idx, row in enumerate(parsed_rows):
                global_index = plan.offset + idx
                if global_index < start_index:
                    continue
                if cancel_event.is_set():
                    break
                aggregated_results.append(row)
                progress.row(global_index, row)
            if cancel_event.is_set():
                break

        progress.done([row.__dict__ for row in rows], aggregated_results)
        return [row.__dict__ for row in rows], aggregated_results

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



