from __future__ import annotations

from dataclasses import dataclass
from typing import List


@dataclass
class RunPlan:
    mode: str  # 'RIH' or 'POOH'
    depth_values: List[float]
    include_pipe_density: bool
    include_force_on_end: bool
    combo_count: int
    offset: int = 0

    @property
    def end_offset(self) -> int:
        return self.offset + self.combo_count


def chunk_depths(depths: List[float], chunk_size: int) -> List[List[float]]:
    if chunk_size <= 0:
        raise ValueError("chunk_size must be positive")
    return [depths[i : i + chunk_size] for i in range(0, len(depths), chunk_size)]
