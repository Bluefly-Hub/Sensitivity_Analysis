from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol


class ProgressCallback(Protocol):
    def __call__(self, event: str, payload: dict[str, Any]) -> None:
        ...


@dataclass
class ProgressReporter:
    emit: ProgressCallback

    def init(self, total_rows: int, **metadata: Any) -> None:
        self.emit("init", {"total_rows": total_rows, **metadata})

    def row(self, index: int, payload: dict[str, Any]) -> None:
        self.emit("row", {"index": index, **payload})

    def done(self, final_inputs: Any, outputs: Any) -> None:
        self.emit("done", {"final_inputs": final_inputs, "outputs": outputs})

