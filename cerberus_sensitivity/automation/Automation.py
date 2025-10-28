from __future__ import annotations

import sys
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from dataclasses import dataclass
from itertools import cycle, product
from typing import Any, Callable, List, Mapping, Protocol, Sequence

import pandas as pd

from cerberus_sensitivity.automation.button_repository import (
    Clear_Value_List,
    Edit_cmdOK,
    Parameter_Matrix_BHA_Depth_Row0,
    Parameter_Matrix_FOE_POOH_Row0,
    Parameter_Matrix_FOE_RIH_Row0,
    Parameter_Matrix_PFD_Row0,
    Parameter_Matrix_Wizard,
    Setup_POOH,
    Populate_Value_List,
    Sensitivity_Analysis_Calculate,
    Sensitivity_Parameter_ok,
    Sensitivity_Table,
    Set_Parameters_RIH,
    button_Sensitivity_Analysis,
)


DEFAULT_MAX_BATCH_SIZE = 200
_VALUE_LIST_CACHE: dict[str, Tuple[float, ...]] = {}

_DEFAULT_TEST_INPUTS: Mapping[str, Sequence[float]] = {
    "density": (8, 9, 10, 11, 12, 12.13),
    "depth": (4553.18, 4355.11, 4177.33, 3744.68, 3705.57, 3540.85, 3500.00),
    "wob_rih": (0, -1350, -1500, 10000.00, -5000),
    "wob_pooh": (0, 1350, 7000, 14000, 18900, 50400, 78000, 93600),
}


def build_test_dataframe(
    overrides: Mapping[str, Sequence[float]] | None = None,
) -> pd.DataFrame:
    """
    Construct a temporary DataFrame for exploratory testing within this module.

    Parameters
    ----------
    overrides:
        Optional mapping that replaces specific default column sequences.
    The resulting frame repeats shorter sequences so columns align, which
    makes it easy to compare a single depth against many densities, for example.
    """
    data = dict(_DEFAULT_TEST_INPUTS)
    if overrides:
        data.update(overrides)
    normalized = _normalize_test_columns(data)
    frame = pd.DataFrame(normalized)
    for column in frame.columns:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")
    return frame


@dataclass(frozen=True)
class ParameterGrid:
    """Container for the unique values that define a sensitivity batch."""

    densities: Sequence[float]
    depths: Sequence[float]
    wobs: Sequence[float]

    @property
    def sample_count(self) -> int:
        return (
            len(self.densities)
            * len(self.depths)
            * len(self.wobs)
        )


@dataclass(frozen=True)
class PlannedBatch:
    """Describes a batch configuration before execution."""

    parameters: Mapping[str, Sequence[float]]
    combinations: List[dict[str, float]]


@dataclass
class BatchResult:
    """Captures the inputs applied to a batch and the resulting table."""

    mode: str
    combinations: List[dict[str, float]]
    parameters: Mapping[str, List[float]]
    table: pd.DataFrame


class ProgressCallback(Protocol):
    def __call__(self, event: str, payload: dict[str, Any]) -> None:
        ...


@dataclass
class ProgressReporter:
    """Lightweight event emitter used by the GUI while automation runs."""

    emit: ProgressCallback

    def init(self, total_rows: int, **metadata: Any) -> None:
        self.emit("init", {"total_rows": total_rows, **metadata})

    def row(self, index: int, payload: dict[str, Any]) -> None:
        self.emit("row", {"index": index, **payload})

    def done(self, final_inputs: Any, outputs: Any) -> None:
        self.emit("done", {"final_inputs": final_inputs, "outputs": outputs})


def run_automation(
    inputs: pd.DataFrame,
    *,
    max_batch_size: int = DEFAULT_MAX_BATCH_SIZE,
) -> dict[str, List[BatchResult]]:
    """
    Entry point used by the GUI and button repository.

    Parameters
    ----------
    inputs:
        DataFrame that contains all sensitivity inputs. The frame is expected to
        include the shared columns `density` and `depth` plus the WOB variants
        `wob_rih` and `wob_pooh`.
    max_batch_size:
        Number of parameter combinations permitted per automation iteration.
    """

    rih_batches = run_rih(inputs, max_batch_size=max_batch_size)
    pooh_batches = run_pooh(inputs, max_batch_size=max_batch_size)

    # Ensure the Sensitivity window is active before driving UI automation.
    button_Sensitivity_Analysis()

    results: dict[str, List[BatchResult]] = {"RIH": [], "POOH": []}

    if rih_batches:
        _prepare_rih_mode()
        for batch_index, planned in enumerate(rih_batches, start=1):
            batch_result = _execute_batch(
                "RIH",
                planned.parameters,
                planned.combinations,
                batch_index,
            )
            results["RIH"].append(batch_result)

    if pooh_batches:
        _prepare_pooh_mode()
        for batch_index, planned in enumerate(pooh_batches, start=1):
            batch_result = _execute_batch(
                "POOH",
                planned.parameters,
                planned.combinations,
                batch_index,
            )
            results["POOH"].append(batch_result)

    _execute_batches("RIH", results["RIH"])
    _execute_batches("POOH", results["POOH"])
    return results


def _prepare_rih_mode() -> None:
    """Ensure the application is configured for RIH batches."""
    Set_Parameters_RIH()


def _prepare_pooh_mode() -> None:
    """Ensure the application is configured for POOH batches."""
    Setup_POOH()


def _plan_chunk_sizes(
    lengths: Mapping[str, int],
    max_batch_size: int,
) -> dict[str, int]:
    if max_batch_size <= 0:
        raise ValueError("max_batch_size must be positive")

    chunk_sizes: dict[str, int] = {name: 1 for name in lengths}
    assigned: set[str] = set()

    for name, length in sorted(lengths.items(), key=lambda item: item[1], reverse=True):
        if length <= 0:
            raise ValueError(f"No values available for parameter '{name}'")

        product_assigned = 1
        for assigned_name in assigned:
            product_assigned *= chunk_sizes[assigned_name]

        allowed = max_batch_size // product_assigned if product_assigned else max_batch_size
        if allowed <= 0:
            allowed = 1

        chunk_sizes[name] = min(length, allowed)
        assigned.add(name)

    return chunk_sizes


def _chunk_sequence(values: Sequence[float], chunk_size: int) -> List[List[float]]:
    if chunk_size <= 0:
        raise ValueError("chunk_size must be positive")
    if len(values) <= chunk_size:
        return [list(values)]
    return [list(values[idx : idx + chunk_size]) for idx in range(0, len(values), chunk_size)]


def _generate_batches_from_grid(
    grid: ParameterGrid,
    max_batch_size: int,
) -> List[PlannedBatch]:
    if grid.sample_count == 0:
        return []

    lengths = {
        "density": len(grid.densities),
        "depth": len(grid.depths),
        "wob": len(grid.wobs),
    }
    chunk_sizes = _plan_chunk_sizes(lengths, max_batch_size)

    density_chunks = _chunk_sequence(list(grid.densities), chunk_sizes["density"])
    depth_chunks = _chunk_sequence(list(grid.depths), chunk_sizes["depth"])
    wob_chunks = _chunk_sequence(list(grid.wobs), chunk_sizes["wob"])

    batches: List[PlannedBatch] = []
    for density_values in density_chunks:
        for depth_values in depth_chunks:
            for wob_values in wob_chunks:
                if not density_values or not depth_values or not wob_values:
                    continue
                combinations = [
                    {"density": density, "depth": depth, "wob": wob}
                    for density, depth, wob in product(density_values, depth_values, wob_values)
                ]
                if not combinations:
                    continue
                parameters = {
                    "density": tuple(density_values),
                    "depth": tuple(depth_values),
                    "wob": tuple(wob_values),
                }
                batches.append(PlannedBatch(parameters=parameters, combinations=combinations))

    return batches


def _execute_batch(
    mode: str,
    parameters: Mapping[str, Sequence[float]],
    combinations: Sequence[dict[str, float]],
    batch_index: int,
) -> BatchResult:
    if not combinations:
        raise ValueError(f"Batch {batch_index} for {mode} is empty.")

    _update_parameter_matrix(mode, parameters)

    Sensitivity_Analysis_Calculate()
    table = Sensitivity_Table().copy()
    table["mode"] = mode
    table["batch_index"] = batch_index

    return BatchResult(
        mode=mode,
        combinations=list(combinations),
        parameters={key: list(value) for key, value in parameters.items()},
        table=table,
    )


def _update_parameter_matrix(
    mode: str,
    parameters: Mapping[str, Sequence[float]],
) -> None:
    Parameter_Matrix_Wizard()

    _apply_value_list("depth", Parameter_Matrix_BHA_Depth_Row0, parameters["depth"])
    _apply_value_list("density", Parameter_Matrix_PFD_Row0, parameters["density"])

    wob_selector = (
        Parameter_Matrix_FOE_RIH_Row0 if mode.upper() == "RIH" else Parameter_Matrix_FOE_POOH_Row0
    )
    wob_key = f"wob_{mode.upper()}"
    _apply_value_list(wob_key, wob_selector, parameters["wob"])

    Sensitivity_Parameter_ok()


def _apply_value_list(
    cache_key: str,
    selector: Callable[[], object],
    values: Sequence[float],
) -> None:
    normalized = tuple(float(value) for value in values)
    if not normalized:
        raise ValueError(f"No values provided for parameter '{cache_key}'.")

    if _VALUE_LIST_CACHE.get(cache_key) == normalized:
        return

    selector()
    Clear_Value_List()
    Populate_Value_List([_format_value(value) for value in values])
    Edit_cmdOK()
    _VALUE_LIST_CACHE[cache_key] = normalized


def _format_value(value: float) -> str:
    if pd.isna(value):
        raise ValueError("Parameter matrix values must not be NaN.")
    return f"{value:.6g}"

def run_rih(
    inputs: pd.DataFrame,
    *,
    max_batch_size: int = DEFAULT_MAX_BATCH_SIZE,
) -> List[PlannedBatch]:
    """Prepare batched parameter combinations for running RIH automation."""
    grid = _build_parameter_grid(
        inputs,
        density_column="density",
        depth_column="depth",
        wob_column="wob_rih",
    )
    return _generate_batches_from_grid(grid, max_batch_size)


def run_pooh(
    inputs: pd.DataFrame,
    *,
    max_batch_size: int = DEFAULT_MAX_BATCH_SIZE,
) -> List[PlannedBatch]:
    """Prepare batched parameter combinations for running POOH automation."""
    grid = _build_parameter_grid(
        inputs,
        density_column="density",
        depth_column="depth",
        wob_column="wob_pooh",
    )
    return _generate_batches_from_grid(grid, max_batch_size)


def _build_parameter_grid(
    inputs: pd.DataFrame,
    *,
    density_column: str,
    depth_column: str,
    wob_column: str,
) -> ParameterGrid:
    _ensure_columns(inputs, density_column, depth_column, wob_column)
    densities = _unique_non_null(inputs[density_column])
    depths = _unique_non_null(inputs[depth_column])
    wobs = _unique_non_null(inputs[wob_column])
    return ParameterGrid(
        densities=densities,
        depths=depths,
        wobs=wobs,
    )

def _execute_batches(mode: str, batches: Sequence[BatchResult]) -> None:
    """
    Placeholder hook that can be wired into the downstream automation flow.
    """
    for batch_index, batch in enumerate(batches, start=1):
        _log_batch_details(mode, batch_index, batch)


def _log_batch_details(mode: str, batch_index: int, batch: BatchResult) -> None:
    print(f"[automation] {mode} batch {batch_index}: {len(batch.combinations)} samples ready")


def _unique_non_null(series: pd.Series) -> List[float]:
    values = series.dropna().unique().tolist()
    if not values:
        raise ValueError(f"No valid values found for column '{series.name}'")
    return values


def _ensure_columns(df: pd.DataFrame, *columns: str) -> None:
    missing = [column for column in columns if column not in df.columns]
    if missing:
        joined = ", ".join(missing)
        raise KeyError(f"Missing required input column(s): {joined}")


def _normalize_test_columns(data: Mapping[str, Sequence[float]]) -> dict[str, List[float]]:
    lengths = [len(values) for values in data.values() if values]
    if not lengths:
        raise ValueError("Provide at least one value for the test DataFrame.")

    max_length = max(lengths)

    normalized: dict[str, List[float]] = {}
    for column, values in data.items():
        if not values:
            normalized[column] = [pd.NA] * max_length
            continue

        if len(values) == max_length:
            normalized[column] = list(values)
        elif len(values) == 1:
            normalized[column] = [values[0]] * max_length
        else:
            iterator = cycle(values)
            normalized[column] = [next(iterator) for _ in range(max_length)]

    return normalized


if __name__ == "__main__":
    test_df = build_test_dataframe()
    summary = run_automation(test_df)
    print("RIH batches:", len(summary["RIH"]))
    print("POOH batches:", len(summary["POOH"]))
