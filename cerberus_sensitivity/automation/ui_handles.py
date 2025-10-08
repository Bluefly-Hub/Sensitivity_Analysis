from __future__ import annotations

import time
from typing import Optional

import uiautomation as auto

from .window_hygiene import WindowHygieneAgent


class UIHandlesAgent:
    def __init__(self, hygiene: WindowHygieneAgent) -> None:
        self.hygiene = hygiene
        self._wizard: Optional[auto.WindowControl] = None

    @property
    def wizard(self) -> auto.WindowControl:
        if self._wizard and self._wizard.Exists(0, 0):
            return self._wizard
        self._wizard = self.hygiene.focus_sensitivity_wizard()
        return self._wizard

    def _wait_for_control(self, builder: callable, timeout: float = 10.0) -> auto.Control:
        deadline = time.time() + timeout
        while time.time() < deadline:
            ctrl = builder()
            if ctrl.Exists(0, 0):
                return ctrl
            time.sleep(0.2)
        raise RuntimeError("UI control not found within timeout.")

    # Buttons
    def open_template_button(self) -> auto.ButtonControl:
        return self._wait_for_control(
            lambda: auto.ButtonControl(searchFromControl=self.wizard, AutomationId="cmdOpen", Name="Open")
        )

    def parameter_matrix_button(self) -> auto.ButtonControl:
        return self._wait_for_control(
            lambda: auto.ButtonControl(searchFromControl=self.wizard, AutomationId="cmdMatrix", Name="Parameter Matrix Wizard...")
        )

    def calculate_button(self) -> auto.ButtonControl:
        return self._wait_for_control(
            lambda: auto.ButtonControl(searchFromControl=self.wizard, AutomationId="cmdCalc", Name="Calculate")
        )

    def ok_button(self) -> auto.ButtonControl:
        return self._wait_for_control(
            lambda: auto.ButtonControl(searchFromControl=self.wizard, Name="OK")
        )

    # Checkboxes on Parameters tab
    def checkbox_rih(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(searchFromControl=self.wizard, AutomationId="chkFOE", Name="RIH")
        )

    def checkbox_pooh(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(searchFromControl=self.wizard, AutomationId="chkFOE_POOH", Name="POOH")
        )

    def checkbox_pipe_density(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(searchFromControl=self.wizard, AutomationId="chkPipeFluidDens", Name="Pipe fluid density")
        )

    def checkbox_depth(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(searchFromControl=self.wizard, AutomationId="chkDepth", Name="BHA depth")
        )

    # Outputs tab checkboxes
    def checkbox_rho_min_surface_weight(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(searchFromControl=self.wizard, AutomationId="chkRIH_MinSW", Name="Minimum surface weight during RIH")
        )

    def checkbox_pooh_max_surface_weight(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(searchFromControl=self.wizard, AutomationId="chkPOOH_MaxSW", Name="Maximum surface weight during POOH")
        )

    def checkbox_pooh_max_pipe_stress(self) -> auto.CheckBoxControl:
        return self._wait_for_control(
            lambda: auto.CheckBoxControl(
                searchFromControl=self.wizard,
                AutomationId="chkPOOH_MaxYield",
                Name="Maximum pipe stress during POOH  (% of YS)",
            )
        )

    # Menus
    def menu_tools(self) -> auto.MenuItemControl:
        def builder() -> auto.MenuItemControl:
            main = self.hygiene.ensure_main_window()
            self.hygiene.bring_forward(main)
            menu = auto.MenuItemControl(searchFromControl=main, Name="Tools")
            if not menu.Exists(0, 0):
                menu = auto.MenuItemControl(searchFromControl=main, NameRegex=r"^\s*Tools\b.*")
            if not menu.Exists(0, 0):
                menu = auto.MenuItemControl(searchDepth=6, Name="Tools")
            return menu

        return self._wait_for_control(builder)

    def menu_sensitivity_analysis(self) -> auto.MenuItemControl:
        tools = self.menu_tools()
        pattern = None
        try:
            pattern = tools.GetExpandCollapsePattern()
        except AttributeError:
            pattern = None
        if pattern is not None:
            pattern.Expand()
        else:
            tools.Click()
        return self._wait_for_control(lambda: auto.MenuItemControl(Name="Sensitivity Analysis..."))

    def menu_file_open_template(self) -> auto.MenuItemControl:
        wizard = self.wizard
        self.hygiene.bring_forward(wizard)
        menu_file = self._wait_for_control(lambda: auto.MenuItemControl(searchFromControl=wizard, Name="File"))
        menu_file.GetInvokePattern().Invoke()
        return self._wait_for_control(lambda: auto.MenuItemControl(Name="Open Template"))

    def menu_edit_copy_tabulated(self) -> auto.MenuItemControl:
        wizard = self.wizard
        self.hygiene.bring_forward(wizard)
        menu_edit = self._wait_for_control(lambda: auto.MenuItemControl(searchFromControl=wizard, Name="Edit"))
        menu_edit.GetInvokePattern().Invoke()
        return self._wait_for_control(lambda: auto.MenuItemControl(Name="Copy tabulated data"))


