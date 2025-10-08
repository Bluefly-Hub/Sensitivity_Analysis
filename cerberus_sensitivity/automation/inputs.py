from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Sequence


@dataclass
class SensitivityInputRow:
    pipe_fluid_density: float | None = None
    sleeve_number: str | None = None
    depth: float | None = None
    stretch_foe_rih: float | None = None
    stretch_foe_pooh: float | None = None


@dataclass
class SensitivityInputs:
    pipe_fluid_densities: List[float] = field(default_factory=list)
    depths: List[float] = field(default_factory=list)
    stretch_foe_rih: List[float] = field(default_factory=list)
    stretch_foe_pooh: List[float] = field(default_factory=list)

    @classmethod
    def from_rows(cls, rows: Sequence[SensitivityInputRow]) -> "SensitivityInputs":
        uniq_density: List[float] = []
        uniq_depths: List[float] = []
        uniq_rih: List[float] = []
        uniq_pooh: List[float] = []

        def append_unique(target: List[float], value: float | None) -> None:
            if value is None:
                return
            if value not in target:
                target.append(value)

        for row in rows:
            append_unique(uniq_density, row.pipe_fluid_density)
            append_unique(uniq_depths, row.depth)
            append_unique(uniq_rih, row.stretch_foe_rih)
            append_unique(uniq_pooh, row.stretch_foe_pooh)

        return cls(
            pipe_fluid_densities=sorted(uniq_density),
            depths=sorted(uniq_depths),
            stretch_foe_rih=sorted(uniq_rih),
            stretch_foe_pooh=sorted(uniq_pooh),
        )
