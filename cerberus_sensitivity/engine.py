from __future__ import annotations

import threading
from typing import Any, Dict, Iterable, Iterator, List, Sequence, Tuple

import pandas as pd

from .automation.Automation import (
    DEFAULT_MAX_BATCH_SIZE,
    BatchResult,
    ProgressReporter,
    run_automation,
)


class CerberusEngine:
    """
    Thin façade used by the GUI to trigger automation batches.

    This class now delegates directly to the streamlined automation skeleton.
    """

    def __init__(self, *, max_batch_size: int | None = None) -> None:
        self.max_batch_size = max_batch_size or DEFAULT_MAX_BATCH_SIZE

    def run_scan(
        self,
        progress: ProgressReporter,
        data_list: Sequence[Any],
        start_index: int,
        cancel_event: threading.Event,
        template_name: str = "auto",
    ) -> Tuple[Sequence[Any], Dict[str, List[BatchResult]]]:
        del template_name  # Template selection is not used in the new skeleton.

        inputs_df = _standardize_inputs(data_list)
        outputs = run_automation(inputs_df, max_batch_size=self.max_batch_size)

        total_samples = _count_samples(outputs)
        progress.init(total_samples, template="skeleton")

        for index, (mode, combo) in enumerate(_iterate_combos(outputs)):
            if cancel_event.is_set():
                break
            if index < start_index:
                continue
            progress.row(
                index,
                {
                    "mode": mode,
                    "density": combo.get("density"),
                    "depth": combo.get("depth"),
                    "wob": combo.get("wob"),
                },
            )

        progress.done(final_inputs=data_list, outputs=outputs)
        return data_list, outputs


def run_scan(
    progress: ProgressReporter,
    data_list: Sequence[Any],
    start_index: int,
    cancel_event: threading.Event,
    template_name: str = "auto",
):
    engine = CerberusEngine()
    return engine.run_scan(progress, data_list, start_index, cancel_event, template_name=template_name)


def _standardize_inputs(rows: Sequence[Any]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=["density", "depth", "wob_rih", "wob_pooh"])

    # Convert row objects to dictionaries when possible.
    normalized_rows: list[dict[str, Any]] = []
    for row in rows:
        if isinstance(row, dict):
            normalized_rows.append(dict(row))
        else:
            attributes = {
                key: getattr(row, key)
                for key in dir(row)
                if not key.startswith("_") and not callable(getattr(row, key))
            }
            normalized_rows.append(attributes)

    frame = pd.DataFrame(normalized_rows)

    column_map = {
        "pipe_fluid_density": "density",
        "depth": "depth",
        "stretch_foe_rih": "wob_rih",
        "stretch_foe_pooh": "wob_pooh",
    }

    data: dict[str, Iterable[Any]] = {}
    for source, target in column_map.items():
        if source in frame.columns:
            data[target] = frame[source]
        else:
            data[target] = [pd.NA] * len(frame)

    inputs_df = pd.DataFrame(data)
    for column in inputs_df.columns:
        inputs_df[column] = pd.to_numeric(inputs_df[column], errors="coerce")

    return inputs_df


def _count_samples(outputs: Dict[str, List[BatchResult]]) -> int:
    return sum(len(batch.combinations) for batches in outputs.values() for batch in batches)


def _iterate_combos(outputs: Dict[str, List[BatchResult]]) -> Iterator[Tuple[str, dict]]:
    for mode, batches in outputs.items():
        for batch in batches:
            for combo in batch.combinations:
                yield mode, combo
