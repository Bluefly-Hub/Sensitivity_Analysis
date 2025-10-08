#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Orpheus UI Button Tester GUI (non-blocking)
-------------------------------------------
- Runs each click in a background thread so the GUI never locks up.
- Uses manual MenuItem clicks (UIA) instead of .menu_select() to avoid blocking UI loops.
- Minimal navigation only; one click per button.
"""
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
import logging
from typing import Optional

from pywinauto import Application, timings
from pywinauto.findwindows import ElementNotFoundError

logger = logging.getLogger("orpheus_gui")
logger.setLevel(logging.INFO)
_console_handler = logging.StreamHandler()
_console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(_console_handler)

MAIN_TITLE_RE = r"^Orpheus - <.*?>\s*$"
WIZ_TITLE_RE  = r"^\s*Sensitivity Analysis Wizard - .*"
OPEN_TPL_RE   = r"^\s*Open Sensitivity Template"
PARAM_WIZ_RE  = r"^\s*Sensitivity Parameter Matrix Wizard"
VAL_EDITOR_RE = r"^\s*Parameter Value Editor - .*"

timings.Timings.window_find_timeout = 12

ACTIONS = {
    # Menus
    "Sensitivity Analysis...":      {"context": "main",          "type": "MenuPath",        "path": ["Tools","Sensitivity Analysis..."]},
    "Open Template":                {"context": "wizard",        "type": "MenuPath",        "path": ["File","Open Template"]},
    "Copy tabulated data":          {"context": "wizard",        "type": "MenuPath",        "path": ["Edit","Copy tabulated data"]},

    # Wizard checkboxes
    "Pipe fluid density":           {"context": "wizard",        "type": "CheckBoxAuto",    "auto_id": "chkPipeFluidDens"},
    "POOH":                         {"context": "wizard",        "type": "CheckBoxAuto",    "auto_id": "chkFOE_POOH"},
    "RIH":                          {"context": "wizard",        "type": "CheckBoxAuto",    "auto_id": "chkFOE"},
    "BHA depth":                    {"context": "wizard",        "type": "CheckBoxAuto",    "auto_id": "chkDepth"},
    "Minimum surface weight during RIH": {"context": "wizard",   "type": "CheckBoxAuto",    "auto_id": "chkRIH_MinSW"},
    "Maximum surface weight during POOH": {"context": "wizard",  "type": "CheckBoxAuto",    "auto_id": "chkPOOH_MaxSW"},
    "Maximum pipe stress during POOH  (% of YS)": {"context": "wizard", "type": "CheckBoxAuto", "auto_id": "chkPOOH_MaxYield"},

    # Tabs/panes
    "Outputs":                      {"context": "wizard",        "type": "OutputsTabOrPane"},
    "Parameters":                   {"context": "wizard",        "type": "TabItem",         "title": "Parameters"},

    # Buttons
    "Parameter Matrix Wizard...":   {"context": "wizard",        "type": "ButtonAuto",      "auto_id": "cmdMatrix"},
    "Calculate":                    {"context": "wizard",        "type": "ButtonAuto",      "auto_id": "cmdCalc"},
    "Open":                         {"context": "open_template", "type": "ButtonAuto",      "auto_id": "cmdOpen"},
    "OK":                           {"context": "value_editor",  "type": "ButtonAuto",      "auto_id": "cmdOK"},
    "(unnamed) Add [cmdAdd]":       {"context": "value_editor",  "type": "ButtonAuto",      "auto_id": "cmdAdd"},
    "(unnamed) Delete [cmdDelete]": {"context": "value_editor",  "type": "ButtonAuto",      "auto_id": "cmdDelete"},

    # Lists/Grid rows
    "auto":                         {"context": "open_template", "type": "ListItemTitle",   "title": "auto"},
    "1":                            {"context": "value_editor",  "type": "ListItemTitle",   "title": "1"},
    "9.58":                         {"context": "value_editor",  "type": "ListItemTitle",   "title": "9.58"},
    "BHA Depth (ft) Row 0":         {"context": "param_wizard",  "type": "DataItemTitleRe", "title_re": r"BHA Depth\s*\(ft\)\s*Row 0"},
    "Pipe Fluid Density (lb/gal) Row 0": {"context": "param_wizard", "type": "DataItemTitleRe", "title_re": r"Pipe Fluid Density\s*\(lb/gal\)\s*Row 0"},
    "Force on End - POOH (lbf) Row 0":   {"context": "param_wizard", "type": "DataItemTitleRe", "title_re": r"Force on End - POOH\s*\(lbf\)\s*Row 0"},
}

def connect_app() -> Application:
    app = Application(backend="uia")
    app.connect(title_re=MAIN_TITLE_RE)
    return app

def wait_visible(win, timeout=12):
    win.wait("visible", timeout=timeout)
    win.wait("enabled", timeout=timeout)
    return win

def get_context_window(app: Application, context: str, ensure_wizard: bool):
    if context == "main":
        return wait_visible(app.window(title_re=MAIN_TITLE_RE))
    if context == "wizard":
        try:
            return wait_visible(app.window(title_re=WIZ_TITLE_RE))
        except ElementNotFoundError:
            if not ensure_wizard:
                raise
            main = get_context_window(app, "main", ensure_wizard=False)
            # Manual non-blocking menu click to open the wizard
            try:
                tools = main.child_window(title="Tools", control_type="MenuItem").wrapper_object()
                tools.click_input()
                time.sleep(0.15)
                itm = main.child_window(title="Sensitivity Analysis...", control_type="MenuItem").wrapper_object()
                itm.click_input()
            except Exception:
                # Fallback to menu_select with small timeout if needed
                try:
                    main.menu_select("Tools->Sensitivity Analysis...")
                except Exception:
                    pass
            return wait_visible(app.window(title_re=WIZ_TITLE_RE))
    if context == "open_template":
        return wait_visible(app.window(title_re=OPEN_TPL_RE))
    if context == "param_wizard":
        return wait_visible(app.window(title_re=PARAM_WIZ_RE))
    if context == "value_editor":
        return wait_visible(app.window(title_re=VAL_EDITOR_RE))
    raise ValueError(f"Unknown context: {context}")

def manual_menu_path_click(win, path_list):
    parent = win
    for label in path_list:
        item = parent.child_window(title=label, control_type="MenuItem").wrapper_object()
        item.click_input()
        time.sleep(0.12)

def click_once(win, spec):
    ttype = spec["type"]
    if ttype == "MenuPath":
        # Use manual non-blocking menu clicks
        manual_menu_path_click(win, spec["path"])
        return True
    if ttype == "ButtonAuto":
        ctrl = win.child_window(auto_id=spec["auto_id"], control_type="Button").wrapper_object()
        ctrl.click_input()
        return True
    if ttype == "CheckBoxAuto":
        ctrl = win.child_window(auto_id=spec["auto_id"], control_type="CheckBox").wrapper_object()
        ctrl.click_input()
        return True
    if ttype == "TabItem":
        try:
            ctrl = win.child_window(title=spec["title"], control_type="TabItem").wrapper_object()
        except Exception:
            ctrl = win.child_window(title=spec["title"], control_type="Pane").wrapper_object()
        ctrl.click_input()
        return True
    if ttype == "ListItemTitle":
        ctrl = win.child_window(title=spec["title"], control_type="ListItem").wrapper_object()
        ctrl.click_input()
        return True
    if ttype == "DataItemTitleRe":
        ctrl = win.child_window(title_re=spec["title_re"], control_type="DataItem").wrapper_object()
        ctrl.click_input()
        return True
    if ttype == "OutputsTabOrPane":
        try:
            ctrl = win.child_window(title="Outputs", control_type="Pane").wrapper_object()
        except Exception:
            ctrl = win.child_window(title="Outputs", control_type="TabItem").wrapper_object()
        ctrl.click_input()
        return True
    raise ValueError(f"Unknown action type: {ttype}")

class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Orpheus UI Button Tester (non-blocking)")
        self.geometry("800x700")

        self.ensure_wizard_var = tk.BooleanVar(value=False)

        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=8, pady=6)

        ttk.Checkbutton(top, text="Open Wizard automatically (Tools → Sensitivity Analysis...)", variable=self.ensure_wizard_var).pack(side=tk.LEFT, padx=(0,10))
        ttk.Button(top, text="Refresh App Connection", command=self.refresh_connection).pack(side=tk.RIGHT, padx=6)
        ttk.Button(top, text="Quit", command=self.destroy).pack(side=tk.RIGHT, padx=6)

        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.frame = ttk.Frame(self.canvas)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((0,0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        groups = [
            ("Main / Wizard Menus", ["Sensitivity Analysis...", "Open Template", "Copy tabulated data"]),
            ("Wizard Checkboxes", ["Parameters","BHA depth", "Pipe fluid density", "RIH", "POOH",
                                   "Minimum surface weight during RIH", "Maximum surface weight during POOH",
                                   "Maximum pipe stress during POOH  (% of YS)", "Outputs"]),
            ("Wizard Buttons", ["Parameter Matrix Wizard...", "Calculate"]),
            ("Open Template Dialog", ["auto", "Open"]),
            ("Parameter Matrix Wizard (Grid Rows)", ["BHA Depth (ft) Row 0",
                                                    "Pipe Fluid Density (lb/gal) Row 0",
                                                    "Force on End - POOH (lbf) Row 0"]),
            ("Value Editor Dialog", ["1", "9.58", "(unnamed) Add [cmdAdd]", "(unnamed) Delete [cmdDelete]", "OK"]),
        ]

        for gtitle, names in groups:
            lf = ttk.LabelFrame(self.frame, text=gtitle)
            lf.pack(fill=tk.X, padx=8, pady=6)
            row = ttk.Frame(lf)
            row.pack(fill=tk.X, padx=6, pady=4)
            for name in names:
                if name not in ACTIONS: continue
                b = ttk.Button(row, text=name, command=lambda n=name: self.dispatch_click(n))
                b.pack(side=tk.LEFT, padx=4, pady=3)

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=False, padx=8, pady=6)
        self.log_text = tk.Text(log_frame, height=10, wrap="word")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self._app = None
        self._busy = False
        self.refresh_connection()

    def set_busy(self, busy: bool):
        self._busy = busy
        self.config(cursor="watch" if busy else "")
        self.after(0, self.update_idletasks)

    def log(self, msg):
        ts = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_text.see(tk.END)
        logger.info(msg)

    def refresh_connection(self):
        try:
            self._app = connect_app()
            self.log("Connected to Orpheus main window.")
        except Exception as e:
            self._app = None
            self.log(f"ERROR: Could not connect to Orpheus: {e}")

    def dispatch_click(self, name: str):
        if self._busy:
            self.log("Busy with a previous click; please wait…")
            return
        self.set_busy(True)
        t = threading.Thread(target=self._click_worker, args=(name,), daemon=True)
        t.start()

    def _click_worker(self, name: str):
        try:
            if self._app is None:
                self.refresh_connection()
                if self._app is None:
                    self._post_log("ERROR: Not connected to Orpheus.")
                    return
            spec = ACTIONS[name]
            ctx = spec["context"]
            # Resolve context (may open wizard non-blocking if requested)
            try:
                win = get_context_window(self._app, ctx, ensure_wizard=self.ensure_wizard_var.get())
            except ElementNotFoundError:
                self._post_log(f"Context window for '{name}' not found (context={ctx}). Open it manually and retry.")
                return
            except Exception as e:
                self._post_log(f"ERROR resolving context for '{name}': {e}")
                return

            try:
                click_once(win, spec)
                self._post_log(f"Clicked '{name}'.")
            except ElementNotFoundError as e:
                self._post_log(f"Element for '{name}' not found in context '{ctx}': {e}")
            except Exception as e:
                self._post_log(f"ERROR clicking '{name}': {e}")
        finally:
            self.after(0, lambda: self.set_busy(False))

    def _post_log(self, msg):
        self.after(0, lambda: self.log(msg))

if __name__ == "__main__":
    AppGUI().mainloop()
