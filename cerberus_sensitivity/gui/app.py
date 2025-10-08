from __future__ import annotations

import csv
import io
import queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Dict, List

from ..automation.progress import ProgressReporter
from ..engine import CerberusEngine


INPUT_COLUMNS = (
    ("pipe_fluid_density", "Density of Pipe Fluid (PPG)"),
    ("sleeve_number", "Input Sleeve"),
    ("depth", "Input Depth (ft)"),
    ("stretch_foe_rih", "RIH-WOB"),
    ("stretch_foe_pooh", "POOH-WOB"),
)


def _normalize_header(value: str) -> str:
    return "".join(ch for ch in value.lower() if ch.isalnum())


class KeepAwake:
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001

    def __enter__(self) -> "KeepAwake":
        self._set_state(self.ES_CONTINUOUS | self.ES_SYSTEM_REQUIRED)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self._set_state(self.ES_CONTINUOUS)

    @staticmethod
    def _set_state(flags: int) -> None:
        try:
            import ctypes

            ctypes.windll.kernel32.SetThreadExecutionState(flags)
        except Exception:
            pass


class CerberusApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Sensitivity Runner")
        self.geometry("1200x700")

        self.engine = CerberusEngine()
        self.event_queue: "queue.Queue[tuple[str, Dict[str, Any]]]" = queue.Queue()
        self.progress = ProgressReporter(self._enqueue_event)
        self.worker_thread: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.total_rows = 0
        self.results_data: List[Dict[str, Any]] = []
        self._keep_awake: KeepAwake | None = None
        self._inputs_enabled = True

        self.input_tree: ttk.Treeview
        self.result_tree: ttk.Treeview
        self.status_var = tk.StringVar(value="Idle")

        self.btn_run: ttk.Button
        self.btn_resume: ttk.Button
        self.btn_stop: ttk.Button
        self.btn_add: ttk.Button
        self.btn_remove: ttk.Button
        self.btn_paste: ttk.Button

        self._build_layout()
        self._set_controls_enabled(True)
        self.after(100, self._process_queue)

    def _build_layout(self) -> None:
        container = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        container.pack(fill=tk.BOTH, expand=True)

        input_frame = ttk.Frame(container, padding=10)
        result_frame = ttk.Frame(container, padding=10)
        container.add(input_frame, weight=2)
        container.add(result_frame, weight=3)

        ttk.Label(input_frame, text="Input Rows").pack(anchor=tk.W)
        self.input_tree = ttk.Treeview(
            input_frame,
            columns=[col for col, _ in INPUT_COLUMNS],
            show="headings",
            selectmode="extended",
        )
        for col, heading in INPUT_COLUMNS:
            self.input_tree.heading(col, text=heading)
            self.input_tree.column(col, width=170, anchor=tk.CENTER)
        self.input_tree.pack(fill=tk.BOTH, expand=True, pady=(5, 5))
        self.input_tree.bind("<Double-1>", self._edit_input_cell)
        self.input_tree.bind("<Delete>", self._delete_selected_rows)

        button_row = ttk.Frame(input_frame)
        button_row.pack(fill=tk.X, pady=5)
        self.btn_add = ttk.Button(button_row, text="Add Row", command=self._add_input_row)
        self.btn_add.pack(side=tk.LEFT)
        self.btn_remove = ttk.Button(button_row, text="Remove Selected", command=self._remove_selected)
        self.btn_remove.pack(side=tk.LEFT, padx=(5, 0))
        self.btn_paste = ttk.Button(button_row, text="Paste Rows", command=self._paste_rows)
        self.btn_paste.pack(side=tk.LEFT, padx=(5, 0))

        control_row = ttk.Frame(input_frame)
        control_row.pack(fill=tk.X, pady=(10, 0))
        self.btn_run = ttk.Button(control_row, text="Run", command=self._run)
        self.btn_run.pack(side=tk.LEFT)
        self.btn_resume = ttk.Button(control_row, text="Resume", command=self._resume)
        self.btn_resume.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(control_row, text="Stop", command=self._stop)
        self.btn_stop.pack(side=tk.LEFT)

        ttk.Label(input_frame, textvariable=self.status_var).pack(anchor=tk.W, pady=(10, 0))

        ttk.Label(result_frame, text="Results").pack(anchor=tk.W)
        self.result_tree = ttk.Treeview(result_frame, columns=(), show="headings", selectmode="browse")
        self.result_tree.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

        export_row = ttk.Frame(result_frame)
        export_row.pack(fill=tk.X)
        ttk.Button(export_row, text="Copy Results", command=self._copy_results).pack(side=tk.LEFT)

    def _set_controls_enabled(self, enabled: bool) -> None:
        self._inputs_enabled = enabled
        buttons = (self.btn_run, self.btn_resume, self.btn_add, self.btn_remove, self.btn_paste)
        if enabled:
            for button in buttons:
                button.state(["!disabled"])
            self.btn_stop.state(["disabled"])
        else:
            for button in buttons:
                button.state(["disabled"])
            self.btn_stop.state(["!disabled"])

    def _add_input_row(self) -> None:
        values = [""] * len(INPUT_COLUMNS)
        self.input_tree.insert("", tk.END, values=values)

    def _remove_selected(self) -> None:
        for item in self.input_tree.selection():
            self.input_tree.delete(item)

    def _delete_selected_rows(self, _event: tk.Event[Any] | None = None) -> None:
        if not self._inputs_enabled:
            return
        self._remove_selected()

    def _edit_input_cell(self, event: tk.Event[Any]) -> None:
        if not self._inputs_enabled:
            return
        region = self.input_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item_id = self.input_tree.identify_row(event.y)
        column_id = self.input_tree.identify_column(event.x)
        if not item_id or not column_id:
            return
        bbox = self.input_tree.bbox(item_id, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        column = self.input_tree.column(column_id, option="id")
        current_value = self.input_tree.set(item_id, column)
        entry = ttk.Entry(self.input_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.focus()

        def finish(_: tk.Event[Any] | None = None) -> None:
            self.input_tree.set(item_id, column, entry.get())
            entry.destroy()

        entry.bind("<FocusOut>", finish)
        entry.bind("<Return>", finish)

    def _collect_inputs(self) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        for item in self.input_tree.get_children():
            values = self.input_tree.item(item, "values")
            row = {key: values[idx] if idx < len(values) else "" for idx, (key, _) in enumerate(INPUT_COLUMNS)}
            rows.append(row)
        return rows

    def _paste_rows(self) -> None:
        if not self._inputs_enabled:
            return
        try:
            raw_text = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Paste Rows", "Clipboard is empty or unavailable.")
            return
        reader = csv.reader(io.StringIO(raw_text), delimiter="\t")
        rows = [[cell.strip() for cell in row] for row in reader if any(cell.strip() for cell in row)]
        if not rows:
            messagebox.showinfo("Paste Rows", "Clipboard does not contain tabular data.")
            return

        normalized_targets = [_normalize_header(title) for _, title in INPUT_COLUMNS]
        normalized_header = [_normalize_header(cell) for cell in rows[0]]

        if all(value in normalized_header for value in normalized_targets):
            header_lookup = {value: normalized_header.index(value) for value in normalized_targets}
            data_rows = rows[1:]
        else:
            header_lookup = {normalized_targets[idx]: idx for idx in range(len(normalized_targets))}
            data_rows = rows

        if not data_rows:
            messagebox.showinfo("Paste Rows", "No data rows detected after the header.")
            return

        for data_row in data_rows:
            values: List[str] = []
            for target in normalized_targets:
                idx = header_lookup.get(target)
                cell_value = data_row[idx] if idx is not None and idx < len(data_row) else ""
                values.append(cell_value)
            if any(value for value in values):
                self.input_tree.insert("", tk.END, values=values)

    def _enqueue_event(self, event: str, payload: Dict[str, Any]) -> None:
        self.event_queue.put((event, payload))

    def _process_queue(self) -> None:
        while True:
            try:
                event, payload = self.event_queue.get_nowait()
            except queue.Empty:
                break
            if event == "init":
                self.total_rows = payload.get("total_rows", 0)
                self.status_var.set(f"Running ({self.total_rows} rows)")
                self._set_controls_enabled(False)
            elif event == "row":
                self._handle_row_event(payload)
            elif event == "done":
                self._handle_done()
            elif event == "error":
                self._handle_error(payload.get("message", "Unexpected error"))
        self.after(100, self._process_queue)

    def _handle_row_event(self, payload: Dict[str, Any]) -> None:
        self.results_data.append(payload)
        if not self.result_tree["columns"]:
            columns = list(payload.keys())
            self.result_tree.configure(columns=columns)
            for col in columns:
                self.result_tree.heading(col, text=col)
                self.result_tree.column(col, width=160, anchor=tk.CENTER)
        values = [payload.get(col, "") for col in self.result_tree["columns"]]
        self.result_tree.insert("", tk.END, values=values)
        processed = len(self.results_data)
        message = f"Processed {processed}/{self.total_rows} rows" if self.total_rows else f"Processed {processed} rows"
        self.status_var.set(message)

    def _handle_done(self) -> None:
        status = "Cancelled" if self.cancel_event.is_set() else "Completed"
        self.status_var.set(status)
        self.cancel_event.clear()
        self._set_controls_enabled(True)
        self.worker_thread = None
        self._keep_awake = None

    def _handle_error(self, message: str) -> None:
        self.status_var.set("Error")
        self._set_controls_enabled(True)
        messagebox.showerror("Automation Error", message)
        self.cancel_event.clear()
        self.worker_thread = None
        self._keep_awake = None

    def _start_worker(self, start_index: int) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Automation", "Worker already running.")
            return
        data_list = self._collect_inputs()
        if not data_list:
            messagebox.showinfo("Automation", "Add at least one input row.")
            return
        if start_index == 0:
            self.results_data.clear()
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            self.result_tree.configure(columns=())
        self.cancel_event.clear()
        keep_awake = KeepAwake()
        self._keep_awake = keep_awake

        def worker() -> None:
            try:
                with keep_awake:
                    self.engine.run_scan(self.progress, data_list, start_index, self.cancel_event)
            except Exception as exc:  # pylint: disable=broad-exception-caught
                self._enqueue_event("error", {"message": str(exc)})

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _run(self) -> None:
        self._start_worker(start_index=0)

    def _resume(self) -> None:
        self._start_worker(start_index=len(self.results_data))

    def _stop(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            self.cancel_event.set()
            self.status_var.set("Stopping...")

    def _copy_results(self) -> None:
        if not self.results_data:
            messagebox.showinfo("Copy Results", "No data to copy yet.")
            return
        columns = list(self.result_tree["columns"])
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=columns, delimiter="\t")
        writer.writeheader()
        for row in self.results_data:
            writer.writerow({key: row.get(key, "") for key in columns})
        self.clipboard_clear()
        self.clipboard_append(output.getvalue())
        messagebox.showinfo("Copy Results", "Results copied to clipboard.")


def launch_gui() -> None:
    app = CerberusApp()
    app.mainloop()


