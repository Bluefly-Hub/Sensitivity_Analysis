from __future__ import annotations

import time
from typing import Iterable

import uiautomation as auto

from .ui_handles import UIHandlesAgent
from .window_hygiene import WindowHygieneAgent


class UITriggerAgent:
    def __init__(self, hygiene: WindowHygieneAgent, handles: UIHandlesAgent) -> None:
        self.hygiene = hygiene
        self.handles = handles

    def open_sensitivity_analysis(self) -> None:
        tools_menu = self.handles.menu_tools()
        tools_menu.Click()
        sensitivity_item = self.handles.menu_sensitivity_analysis()
        sensitivity_item.Click()
        self.hygiene.focus_sensitivity_wizard()

    def load_template(self, template_name: str, timeout: float = 10.0) -> None:
        menu = self.handles.menu_file_open_template()
        menu.Click()
        list_control = auto.ListControl(searchFromControl=self.handles.wizard)
        deadline = time.time() + timeout
        while time.time() < deadline:
            item = auto.ListItemControl(searchFromControl=list_control, Name=template_name)
            if item.Exists(0, 0):
                item.Click()
                break
            time.sleep(0.2)
        else:
            raise RuntimeError(f"Template '{template_name}' not located.")
        self.handles.open_template_button().Click()
        self.hygiene.focus_sensitivity_wizard()

    def configure_for_mode(self, mode: str) -> None:
        """Ensure the Parameters tab is active and configure the mode toggles."""
        self._select_tab("Parameters")
        if mode == "RIH":
            self._set_checkbox(self.handles.checkbox_rih(), True)
            self._set_checkbox(self.handles.checkbox_pooh(), False)
            self._set_checkbox(self.handles.checkbox_pipe_density(), True)
            self._set_checkbox(self.handles.checkbox_depth(), True)
        elif mode == "POOH":
            self._set_checkbox(self.handles.checkbox_rih(), False)
            self._set_checkbox(self.handles.checkbox_pooh(), True)
            self._set_checkbox(self.handles.checkbox_pipe_density(), True)
            self._set_checkbox(self.handles.checkbox_depth(), True)
        else:
            raise ValueError(f"Unsupported mode '{mode}'.")

    def configure_outputs_for_mode(self, mode: str) -> None:
        """Toggle the Outputs tab checkboxes for the selected mode."""
        self._select_tab("Outputs")
        if mode == "RIH":
            self._set_checkbox(self.handles.checkbox_rho_min_surface_weight(), True)
            self._set_checkbox(self.handles.checkbox_pooh_max_surface_weight(), False)
            self._set_checkbox(self.handles.checkbox_pooh_max_pipe_stress(), False)
        else:
            self._set_checkbox(self.handles.checkbox_rho_min_surface_weight(), False)
            self._set_checkbox(self.handles.checkbox_pooh_max_surface_weight(), True)
            self._set_checkbox(self.handles.checkbox_pooh_max_pipe_stress(), True)
        self._select_tab("Parameters")

    def update_parameter_values(
        self,
        parameter_caption: str,
        values: Iterable[float],
        ensure_clear: bool = True,
    ) -> None:
        wizard = self.handles.wizard
        grid = auto.TableControl(searchFromControl=wizard)
        row = auto.DataItemControl(searchFromControl=grid, NameRegex=fr"^{parameter_caption}.*")
        if not row.Exists(0, 0):
            raise RuntimeError(f"Parameter row '{parameter_caption}' not found in matrix.")
        row.DoubleClick()
        editor = auto.WindowControl(NameRegex=fr"^.*Parameter Value Editor - {parameter_caption}.*")
        if not editor.Exists(0, 0):
            raise RuntimeError(f"Value editor for '{parameter_caption}' did not open.")
        list_control = auto.ListControl(searchFromControl=editor)
        if ensure_clear:
            for item in list(list_control.GetChildren())[::-1]:
                if isinstance(item, auto.ListItemControl) and item.Exists(0, 0):
                    item.Select()
                    auto.ButtonControl(searchFromControl=editor, Name="Remove Value").Click()
        add_button = auto.ButtonControl(searchFromControl=editor, Name="Add Value")
        value_field = auto.EditControl(searchFromControl=editor)
        for raw_value in values:
            raw_text = str(raw_value)
            if not raw_text:
                continue
            self._set_edit_value(value_field, raw_text)
            add_button.Click()
        auto.ButtonControl(searchFromControl=editor, Name="OK").Click()
        self.hygiene.bring_forward(editor)

    def open_parameter_matrix(self) -> None:
        self.handles.parameter_matrix_button().Click()
        matrix_window = auto.WindowControl(NameRegex=r".*Sensitivity Parameter Matrix Wizard.*")
        if not matrix_window.Exists(3, 0):
            raise RuntimeError("Parameter Matrix Wizard did not open.")
        matrix_window.SetFocus()

    def close_parameter_matrix(self) -> None:
        matrix_window = auto.WindowControl(NameRegex=r".*Sensitivity Parameter Matrix Wizard.*")
        if matrix_window.Exists(0, 0):
            auto.ButtonControl(searchFromControl=matrix_window, Name="OK").Click()

    def recalc_and_copy_results(self, clipboard_copy_timeout: float = 60.0) -> str:
        calc_button = self.handles.calculate_button()
        calc_button.Click()
        self._wait_for_idle(1.0)
        deadline = time.time() + clipboard_copy_timeout
        while time.time() < deadline:
            if not calc_button.IsEnabled:
                time.sleep(0.2)
                continue
            break
        copy_item = self.handles.menu_edit_copy_tabulated()
        auto.SetClipboardText("")
        copy_item.GetInvokePattern().Invoke()
        self._wait_for_idle(2.0)
        text = auto.GetClipboardText()
        if not text:
            raise RuntimeError("Clipboard did not receive tabulated data.")
        return text

    def _select_tab(self, name: str) -> auto.TabItemControl:
        wizard = self.handles.wizard
        tab = auto.TabItemControl(searchFromControl=wizard, Name=name)
        if not tab.Exists(0, 0):
            raise RuntimeError(f"Tab '{name}' not found.")
        pattern = None
        try:
            pattern = tab.GetSelectionItemPattern()
        except AttributeError:
            pattern = None
        if pattern is not None:
            pattern.Select()
        elif hasattr(tab, "Select"):
            tab.Select()
        else:
            tab.Click()
        return tab

    def _set_edit_value(self, edit: auto.EditControl, value: str) -> None:
        if not value:
            return
        pattern = None
        try:
            pattern = edit.GetValuePattern()
        except AttributeError:
            pattern = None
        if pattern is not None:
            pattern.SetValue(value)
            return
        edit.Click()
        auto.SendKeys(value, interval=0.01)

    def _wait_for_idle(self, seconds: float) -> None:
        try:
            auto.WaitForIdle(seconds)
        except AttributeError:
            time.sleep(seconds)

    def _set_checkbox(self, checkbox: auto.CheckBoxControl, desired_state: bool) -> None:
        current_state = checkbox.GetToggleState()
        if current_state == desired_state:
            return
        pattern = checkbox.GetTogglePattern()
        if pattern is not None:
            if desired_state:
                pattern.SetOn()
            else:
                pattern.SetOff()
        else:
            checkbox.Click()
