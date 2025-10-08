from __future__ import annotations

import argparse
import sys
import time
from dataclasses import dataclass
from typing import Callable, Optional

import uiautomation as auto

from .automation.ui_handles import UIHandlesAgent
from .automation.ui_trigger import UITriggerAgent
from .automation.window_hygiene import WindowHygieneAgent


@dataclass
class ProbeContext:
    hygiene: WindowHygieneAgent
    handles: UIHandlesAgent
    trigger: UITriggerAgent


def build_context(window_regex: str | None = None, timeout: float = 20.0) -> ProbeContext:
    hygiene = WindowHygieneAgent(window_regex=window_regex or WindowHygieneAgent().window_regex, timeout=timeout)
    handles = UIHandlesAgent(hygiene)
    trigger = UITriggerAgent(hygiene, handles)
    return ProbeContext(hygiene=hygiene, handles=handles, trigger=trigger)


Action = Callable[[ProbeContext, argparse.Namespace], Optional[auto.Control]]


def action_focus_main(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    control = ctx.hygiene.ensure_main_window()
    ctx.hygiene.bring_forward(control)
    return control


def action_focus_wizard(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    control = ctx.hygiene.focus_sensitivity_wizard()
    ctx.hygiene.bring_forward(control)
    return control


def action_open_tools_menu(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    menu = ctx.handles.menu_tools()
    menu.Click()
    return menu


def action_open_sensitivity_menu(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    menu = ctx.handles.menu_sensitivity_analysis()
    menu.Click()
    return menu


def action_open_template_button(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    if args.via_menu:
        action_open_sensitivity_menu(ctx, args)
        ctx.handles.menu_file_open_template().Click()
    button = ctx.handles.open_template_button()
    button.Click()
    return button


def action_parameter_matrix(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    button = ctx.handles.parameter_matrix_button()
    button.Click()
    return button


def action_calculate(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    button = ctx.handles.calculate_button()
    button.Click()
    return button


def _toggle_checkbox(control: auto.CheckBoxControl, desired: Optional[bool]) -> None:
    if desired is None:
        control.Click()
        return
    current = control.GetTogglePattern().CurrentToggleState if control.GetTogglePattern() else control.GetToggleState()
    is_checked = current == auto.ToggleState.On if current is not None else False
    if is_checked != desired:
        control.Click()


def action_checkbox(ctx: ProbeContext, args: argparse.Namespace) -> auto.Control:
    mapping = {
        "rih": ctx.handles.checkbox_rih,
        "pooh": ctx.handles.checkbox_pooh,
        "pipe_density": ctx.handles.checkbox_pipe_density,
        "depth": ctx.handles.checkbox_depth,
        "rih_min_sw": ctx.handles.checkbox_rho_min_surface_weight,
        "pooh_max_sw": ctx.handles.checkbox_pooh_max_surface_weight,
        "pooh_max_pipe_stress": ctx.handles.checkbox_pooh_max_pipe_stress,
    }
    if args.name not in mapping:
        raise ValueError(f"Unknown checkbox key '{args.name}'. Available: {', '.join(mapping)}")
    control = mapping[args.name]()
    desired_state = None
    if args.state is not None:
        desired_state = args.state.lower() in {"1", "true", "on", "yes"}
    _toggle_checkbox(control, desired_state)
    return control


def action_wait(ctx: ProbeContext, args: argparse.Namespace) -> None:
    time.sleep(args.seconds)
    return None


ACTIONS: dict[str, Action] = {
    "focus-main": action_focus_main,
    "focus-wizard": action_focus_wizard,
    "menu-tools": action_open_tools_menu,
    "menu-sensitivity": action_open_sensitivity_menu,
    "open-template": action_open_template_button,
    "parameter-matrix": action_parameter_matrix,
    "calculate": action_calculate,
    "checkbox": action_checkbox,
    "sleep": action_wait,
}


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Manual UI probe utilities for Cerberus automation")
    parser.add_argument("action", choices=sorted(ACTIONS.keys()), help="Which action to run")
    parser.add_argument("--state", dest="state", help="Target state for checkbox (true/false) when using action 'checkbox'")
    parser.add_argument("--name", dest="name", help="Checkbox key when using action 'checkbox'")
    parser.add_argument("--seconds", dest="seconds", type=float, default=1.0, help="Sleep duration for 'sleep' action")
    parser.add_argument("--window-regex", dest="window_regex", help="Override the main window regex if needed")
    parser.add_argument("--timeout", dest="timeout", type=float, default=20.0, help="Timeout for locating windows")
    parser.add_argument("--via-menu", dest="via_menu", action="store_true", help="For open-template, trigger File -> Open Template first")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    ctx = build_context(window_regex=args.window_regex, timeout=args.timeout)
    action = ACTIONS[args.action]
    try:
        control = action(ctx, args)
        if isinstance(control, auto.Control):
            control.SetFocus()
            print(f"Action '{args.action}' succeeded on control: {control.Name}")
        else:
            print(f"Action '{args.action}' completed.")
        return 0
    except Exception as exc:  # pragma: no cover - manual utility
        print(f"Action '{args.action}' failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
