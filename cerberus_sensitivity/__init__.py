from __future__ import annotations

from typing import Any

__all__ = ["CerberusEngine", "run_scan", "launch_gui"]


def __getattr__(name: str) -> Any:
    if name in {"run_scan", "CerberusEngine"}:
        from .engine import run_scan, CerberusEngine  # pylint: disable=import-outside-toplevel

        if name == "run_scan":
            return run_scan
        return CerberusEngine
    if name == "launch_gui":
        from .gui.app import launch_gui  # pylint: disable=import-outside-toplevel

        return launch_gui
    raise AttributeError(f"module 'cerberus_sensitivity' has no attribute '{name}'")
