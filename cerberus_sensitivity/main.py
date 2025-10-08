from __future__ import annotations

import os
import sys

if __package__ in (None, ""):
    package_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if package_root not in sys.path:
        sys.path.insert(0, package_root)

from cerberus_sensitivity.engine import run_scan  # pragma: no cover
from cerberus_sensitivity.gui.app import launch_gui  # pragma: no cover

__all__ = ["run_scan", "launch_gui"]


if __name__ == "__main__":
    launch_gui()
