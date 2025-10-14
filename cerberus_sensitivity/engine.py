from __future__ import annotations

import threading
from typing import Any, Sequence

from .automation.orchestrator import CerberusOrchestrator
from .automation.progress import ProgressReporter


class CerberusEngine:
    def __init__(self) -> None:
        self.orchestrator = CerberusOrchestrator()

    def run_scan(
        self,
        progress: ProgressReporter,
        data_list: Sequence[Any],
        start_index: int,
        cancel_event: threading.Event,
        template_name: str = "auto",
    ):
        return self.orchestrator.run_scan(progress, data_list, start_index, cancel_event, template_name=template_name)


def run_scan(progress: ProgressReporter, data_list: Sequence[Any], start_index: int, cancel_event: threading.Event, template_name: str = "auto"):
    engine = CerberusEngine()
    return engine.run_scan(progress, data_list, start_index, cancel_event, template_name=template_name)
