from __future__ import annotations

import ctypes
import time
from typing import Optional

import uiautomation as auto

CERBERUS_WINDOW_REGEX = ".*(Cerberus|Orpheus).*"
SENSITIVITY_WINDOW_REGEX = ".*Sensitivity Analysis Wizard.*"


class WindowHygieneAgent:
    def __init__(self, window_regex: str = CERBERUS_WINDOW_REGEX, timeout: float = 15.0) -> None:
        self.window_regex = window_regex
        self.timeout = timeout
        self._cached_window: Optional[auto.WindowControl] = None

    def ensure_main_window(self) -> auto.WindowControl:
        window = self._cached_window
        if window and window.Exists(0, 0):
            return window
        deadline = time.time() + self.timeout
        search_targets = [
            {"ProcessName": "Orpheus"},
            {"AutomationId": "frmOrpheus"},
            {"NameRegex": self.window_regex},
            {"NameRegex": r".*Orpheus.*"},
            {"NameRegex": r".*Cerberus.*"},
        ]
        while time.time() < deadline:
            for target in search_targets:
                candidate = auto.WindowControl(searchDepth=8, **target)
                if candidate.Exists(0, 0):
                    self._cached_window = candidate
                    return candidate
            time.sleep(0.2)
        raise RuntimeError("Unable to locate the Cerberus/Orpheus main window.")

    def bring_forward(self, control: Optional[auto.Control] = None) -> None:
        wnd = control if control is not None else self.ensure_main_window()
        if isinstance(wnd, auto.Control):
            wnd.SetFocus()
            if hasattr(wnd, "GetWindowPattern"):
                pattern = wnd.GetWindowPattern()
                if pattern is not None:
                    pattern.SetWindowVisualState(auto.WindowVisualState.Normal)
                    pattern.SetWindowVisualState(auto.WindowVisualState.Maximized)
                    pattern.SetWindowVisualState(auto.WindowVisualState.Normal)

    def focus_sensitivity_wizard(self) -> auto.WindowControl:
        main = self.ensure_main_window()
        wizard = auto.WindowControl(searchFromControl=main, NameRegex=SENSITIVITY_WINDOW_REGEX)
        if not wizard.Exists(0, 0):
            wizard = auto.WindowControl(NameRegex=SENSITIVITY_WINDOW_REGEX)
        if not wizard.Exists(0, 0):
            raise RuntimeError("Sensitivity Analysis Wizard window is not open.")
        self.bring_forward(wizard)
        return wizard






