from __future__ import annotations

import csv
import io
import queue
import threading
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Dict, List

import pandas as pd

from Automation import BatchResult, ProgressReporter, CerberusEngine
from clear_comtypes_cache import clear_cache



INPUT_COLUMNS = (
    ("pipe_fluid_density", "Density of Pipe Fluid (PPG)"),
    ("sleeve_number", "Input Sleeve"),
    ("depth", "Input Depth (ft)"),
    ("stretch_foe_rih", "RIH-WOB"),
    ("stretch_foe_pooh", "POOH-WOB"),
)

MODE_RIH = "RIH"
MODE_POOH = "POOH"


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
        # Step 1: Configure base window and automation handles
        self.title("Sensitivity Runner")
        self.geometry("1200x720")

        self.log_path = Path(__file__).resolve().parents[2] / "logs" / "automation_errors.log"
        self.log_path.parent.mkdir(parents=True, exist_ok=True)

        self.engine = CerberusEngine()
        self.event_queue: queue.Queue = queue.Queue()
        self.progress = ProgressReporter(self._enqueue_event)
        self.worker_thread: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.total_rows = 0
        self.status_var = tk.StringVar(value="Idle")
        self.timer_var = tk.StringVar(value="Elapsed: 00:00:00")
        self._timer_start: float | None = None
        self._timer_job: str | None = None
        self._keep_awake: KeepAwake | None = None
        self._inputs_enabled = True

        # Step 1b: Result caches for per-mode aggregation
        self.rih_rows: List[Dict[str, Any]] = []
        self.pooh_rows: List[Dict[str, Any]] = []
        self.rih_df = pd.DataFrame()
        self.pooh_df = pd.DataFrame()
        self._rih_columns: List[str] = []
        self._pooh_columns: List[str] = []
        self._progress_rows: Dict[str, List[Dict[str, Any]]] = {MODE_RIH: [], MODE_POOH: []}
        self._progress_columns: Dict[str, List[str]] = {MODE_RIH: [], MODE_POOH: []}
        self._processed_counts: Dict[str, int] = {MODE_RIH: 0, MODE_POOH: 0}

        # Step 2: UI widget placeholders
        self.input_tree: ttk.Treeview
        self.rih_tree: ttk.Treeview
        self.pooh_tree: ttk.Treeview
        self.btn_run: ttk.Button
        self.btn_stop: ttk.Button
        self.btn_add: ttk.Button
        self.btn_remove: ttk.Button
        self.btn_paste: ttk.Button
        self.btn_copy_rih: ttk.Button
        self.btn_copy_pooh: ttk.Button

        self._build_layout()
        self._set_controls_enabled(True)
        self.after(100, self._process_queue)

    def _build_layout(self) -> None:
        # Step 2: Build input capture panel
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

        # Step 3: Execution controls and status updates
        control_row = ttk.Frame(input_frame)
        control_row.pack(fill=tk.X, pady=(10, 0))
        self.btn_run = ttk.Button(control_row, text="Run", command=self._run)
        self.btn_run.pack(side=tk.LEFT)
        self.btn_stop = ttk.Button(control_row, text="Stop", command=self._stop)
        self.btn_stop.pack(side=tk.LEFT)

        ttk.Label(input_frame, textvariable=self.status_var).pack(anchor=tk.W, pady=(10, 0))
        ttk.Label(input_frame, textvariable=self.timer_var).pack(anchor=tk.W)

        # Step 4: Output panes for RIH/POOH review
        ttk.Label(result_frame, text="Results").pack(anchor=tk.W)
        notebook = ttk.Notebook(result_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(5, 5))

        rih_tab = ttk.Frame(notebook)
        pooh_tab = ttk.Frame(notebook)
        notebook.add(rih_tab, text="RIH Results")
        notebook.add(pooh_tab, text="POOH Results")

        self.rih_tree = self._create_result_tree(rih_tab)
        self.pooh_tree = self._create_result_tree(pooh_tab)

        copy_row = ttk.Frame(result_frame)
        copy_row.pack(fill=tk.X)
        self.btn_copy_rih = ttk.Button(copy_row, text="Copy RIH Results", command=lambda: self._copy_results(MODE_RIH))
        self.btn_copy_rih.pack(side=tk.LEFT)
        self.btn_copy_pooh = ttk.Button(copy_row, text="Copy POOH Results", command=lambda: self._copy_results(MODE_POOH))
        self.btn_copy_pooh.pack(side=tk.LEFT, padx=(5, 0))

    def _create_result_tree(self, parent: ttk.Frame) -> ttk.Treeview:
        tree = ttk.Treeview(parent, columns=(), show="headings", selectmode="browse")
        tree.pack(fill=tk.BOTH, expand=True)
        return tree

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        for widget in (
            self.btn_run,
            self.btn_stop,
            self.btn_add,
            self.btn_remove,
            self.btn_paste,
            self.btn_copy_rih,
            self.btn_copy_pooh,
        ):
            widget.configure(state=state)
        self._inputs_enabled = enabled

    def _collect_inputs(self) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        for item in self.input_tree.get_children():
            values = self.input_tree.item(item, "values")
            row = {col: values[idx] for idx, (col, _) in enumerate(INPUT_COLUMNS)}
            if any(value not in ("", None) for value in row.values()):
                rows.append(row)
        return rows

    def _add_input_row(self) -> None:
        self.input_tree.insert("", tk.END, values=[""] * len(INPUT_COLUMNS))

    def _remove_selected(self) -> None:
        for item in self.input_tree.selection():
            self.input_tree.delete(item)

    def _edit_input_cell(self, event: tk.Event) -> None:  # type: ignore[override]
        if not self._inputs_enabled:
            return
        region = self.input_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.input_tree.identify_row(event.y)
        column_id = self.input_tree.identify_column(event.x)
        if not row_id or not column_id:
            return
        bbox = self.input_tree.bbox(row_id, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        column_index = int(column_id[1:]) - 1
        current_value = self.input_tree.set(row_id, column_index)

        entry = ttk.Entry(self.input_tree)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        entry.place(x=x, y=y, width=width, height=height)
        entry.bind("<Return>", lambda e: self._finish_edit_input(entry, row_id, column_index))
        entry.bind("<FocusOut>", lambda e: self._finish_edit_input(entry, row_id, column_index))

    def _finish_edit_input(self, entry: ttk.Entry, row_id: str, column_index: int) -> None:
        new_value = entry.get()
        entry.destroy()
        self.input_tree.set(row_id, column_index, new_value)

    def _delete_selected_rows(self, event: tk.Event) -> None:  # type: ignore[override]
        if not self._inputs_enabled:
            return
        self._remove_selected()

    def _paste_rows(self) -> None:
        try:
            raw = self.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("Paste Rows", "Clipboard does not contain text data.")
            return
        rows = list(self._parse_clipboard_rows(raw))
        if not rows:
            messagebox.showinfo("Paste Rows", "No tabular rows detected in the clipboard.")
            return
        for row in rows:
            values = [row.get(col, "") for col, _ in INPUT_COLUMNS]
            self.input_tree.insert("", tk.END, values=values)

    def _parse_clipboard_rows(self, raw: str) -> List[Dict[str, Any]]:
        if not raw:
            return []

        reader = csv.reader(io.StringIO(raw), delimiter="\t")
        rows = [list(row) for row in reader]
        if not rows:
            return []

        normalized_targets = [_normalize_header(value) for _, value in INPUT_COLUMNS]
        normalized_sources = [_normalize_header(value) for value in rows[0]]

        header_lookup: Dict[str, int] = {}
        data_rows = rows
        if rows and normalized_sources:
            has_header = all(target in normalized_sources for target in normalized_targets)
            if has_header:
                header_lookup = {value: idx for idx, value in enumerate(normalized_sources)}
                data_rows = rows[1:]

        if not header_lookup:
            header_lookup = {target: idx for idx, target in enumerate(normalized_targets)}

        parsed: List[Dict[str, Any]] = []
        for raw_row in data_rows:
            row_map: Dict[str, Any] = {}
            non_empty = False
            for column_index, (target, (dest_key, _)) in enumerate(zip(normalized_targets, INPUT_COLUMNS)):
                idx = header_lookup.get(target, column_index)
                cell_value = raw_row[idx] if idx < len(raw_row) else ""
                if isinstance(cell_value, str):
                    cell_value = cell_value.strip()
                row_map[dest_key] = cell_value
                if cell_value not in ("", None):
                    non_empty = True
            if non_empty:
                parsed.append(row_map)
        return parsed

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
                template = payload.get("template")
                status = f"Running ({self.total_rows} rows)"
                if template:
                    status += f" | Template: {template}"
                self.status_var.set(status)
                self._set_controls_enabled(False)
                for mode in (MODE_RIH, MODE_POOH):
                    self._progress_rows[mode].clear()
                    self._progress_columns[mode].clear()
                    self._processed_counts[mode] = 0
                    tree = self.pooh_tree if mode == MODE_POOH else self.rih_tree
                    for item in tree.get_children():
                        tree.delete(item)
                    tree.configure(columns=())
            elif event == "row":
                self._handle_row_event(payload)
            elif event == "done":
                self._handle_done(payload)
            elif event == "error":
                self._handle_error(payload.get("message", "Unexpected error"))
        self.after(100, self._process_queue)

    def _handle_row_event(self, payload: Dict[str, Any]) -> None:
        mode = (payload.get("mode") or MODE_RIH).upper()
        row_data = dict(payload)
        row_data.pop("mode", None)
        self._processed_counts[mode] = self._processed_counts.get(mode, 0) + 1

        progress_rows = self._progress_rows[mode]
        progress_columns = self._progress_columns[mode]
        tree = self.pooh_tree if mode == MODE_POOH else self.rih_tree

        progress_rows.append(row_data)
        if not progress_columns:
            progress_columns.extend(row_data.keys())
        else:
            for key in row_data.keys():
                if key not in progress_columns:
                    progress_columns.append(key)
        tree.configure(columns=progress_columns)
        for col in progress_columns:
            tree.heading(col, text=col)
            tree.column(col, width=160, anchor=tk.CENTER)
        values = [row_data.get(col, "") for col in progress_columns]
        tree.insert("", tk.END, values=values)
        processed = self._total_processed_rows()
        message = (
            f"Processed {processed}/{self.total_rows} rows"
            if self.total_rows
            else f"Processed {processed} rows"
        )
        self.status_var.set(message)

    def _handle_done(self, payload: Dict[str, Any]) -> None:
        self._stop_timer()
        outputs = payload.get("outputs")
        if isinstance(outputs, dict):
            self._sync_outputs_from_payload(outputs)
        status = "Cancelled" if self.cancel_event.is_set() else "Completed"
        self.status_var.set(status)
        self.cancel_event.clear()
        self._set_controls_enabled(True)
        self.worker_thread = None
        self._keep_awake = None

    def _sync_outputs_from_payload(self, outputs: Dict[str, Any]) -> None:
        updated_modes: List[str] = []
        for mode_key, rows in outputs.items():
            normalized = mode_key.upper()
            if isinstance(rows, pd.DataFrame):
                incoming = rows.to_dict(orient="records")
            elif isinstance(rows, list):
                if rows and all(isinstance(item, BatchResult) for item in rows):
                    incoming: List[Dict[str, Any]] = []
                    for batch in rows:
                        table = batch.table if isinstance(batch.table, pd.DataFrame) else None
                        if table is None:
                            continue
                        batch_records = table.to_dict(orient="records")
                        incoming.extend(batch_records)
                else:
                    incoming = []
                    for item in rows:
                        if isinstance(item, dict):
                            incoming.append(dict(item))
                        elif hasattr(item, "_asdict"):
                            incoming.append(dict(item._asdict()))  # type: ignore[call-arg]
                        else:
                            incoming.append({"value": item})
            else:
                continue
            target_rows, _, columns = self._resolve_mode_targets(normalized)
            target_rows.clear()
            if incoming:
                target_rows.extend(dict(row) for row in incoming)
                columns.clear()
                columns.extend(incoming[0].keys())
            else:
                columns.clear()
            self._progress_rows[normalized].clear()
            self._progress_columns[normalized].clear()
            updated_modes.append(normalized)

        self.rih_df = pd.DataFrame(self.rih_rows)
        self.pooh_df = pd.DataFrame(self.pooh_rows)
        self._rih_columns = list(self.rih_df.columns)
        self._pooh_columns = list(self.pooh_df.columns)

        for mode in updated_modes:
            self._reload_tree(mode)

    def _handle_error(self, message: str) -> None:
        self.status_var.set("Error")
        self._stop_timer()
        self._set_controls_enabled(True)
        self._log_error(message)
        messagebox.showerror("Automation Error", message)
        self.cancel_event.clear()
        self.worker_thread = None
        self._keep_awake = None

    def _log_error(self, message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message.strip()}\n"
        try:
            with self.log_path.open("a", encoding="utf-8") as log_file:
                log_file.write(entry)
                log_file.write("\n")
        except Exception:
            # Logging should never interrupt the UI flow.
            pass

    def _start_worker(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Automation", "Worker already running.")
            return
        data_list = self._collect_inputs()
        if not data_list:
            messagebox.showinfo("Automation", "Add at least one input row.")
            return
        self._clear_results()
        self._start_timer()
        self.cancel_event.clear()
        keep_awake = KeepAwake()
        self._keep_awake = keep_awake

        def worker() -> None:
            try:
                with keep_awake:
                    self.engine.run_scan(self.progress, data_list, 0, self.cancel_event)
            except Exception as exc:  # pylint: disable=broad-exception-caught
                self._enqueue_event("error", {"message": str(exc)})

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def _run(self) -> None:
        clear_cache()
        self._start_worker()

    def _stop(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            self.cancel_event.set()
            self.status_var.set("Stopping...")

    def _copy_results(self, mode: str) -> None:
        rows, _, _ = self._resolve_mode_targets(mode)
        df = self.rih_df if mode.upper() == MODE_RIH else self.pooh_df
        if df.empty:
            df = pd.DataFrame(rows)
        if df.empty:
            messagebox.showinfo("Copy Results", f"No {mode} results to copy yet.")
            return
        drop_columns = [col for col in ("mode", "batch_index") if col in df.columns]
        export_df = df.drop(columns=drop_columns, errors="ignore")
        export_df = export_df.fillna("")
        export_df = export_df.astype(str)
        buffer = io.StringIO()
        export_df.to_csv(buffer, sep="\t", index=False, lineterminator="\n")
        self.clipboard_clear()
        self.clipboard_append(buffer.getvalue())
        messagebox.showinfo("Copy Results", f"{mode} results copied to clipboard.")

    def _resolve_mode_targets(self, mode: str):
        normalized = mode.upper()
        if normalized == MODE_POOH:
            return self.pooh_rows, self.pooh_tree, self._pooh_columns
        return self.rih_rows, self.rih_tree, self._rih_columns

    def _total_processed_rows(self) -> int:
        return sum(self._processed_counts.values())

    def _clear_results(self) -> None:
        self._reset_timer_display()
        self.rih_rows.clear()
        self.pooh_rows.clear()
        self.rih_df = pd.DataFrame()
        self.pooh_df = pd.DataFrame()
        self._rih_columns.clear()
        self._pooh_columns.clear()
        for mode in (MODE_RIH, MODE_POOH):
            self._progress_rows[mode].clear()
            self._progress_columns[mode].clear()
            self._processed_counts[mode] = 0
        for tree in (self.rih_tree, self.pooh_tree):
            for item in tree.get_children():
                tree.delete(item)
            tree.configure(columns=())

    def _reload_tree(self, mode: str) -> None:
        rows, tree, columns = self._resolve_mode_targets(mode)
        for item in tree.get_children():
            tree.delete(item)
        if not rows:
            tree.configure(columns=())
            return
        if not columns:
            columns.extend(rows[0].keys())
        tree.configure(columns=columns)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=160, anchor=tk.CENTER)
        for row in rows:
            values = [row.get(col, "") for col in columns]
            tree.insert("", tk.END, values=values)

    @staticmethod
    def _format_elapsed(seconds: float) -> str:
        total_seconds = int(seconds)
        hours, remainder = divmod(total_seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours:02}:{minutes:02}:{secs:02}"

    def _start_timer(self) -> None:
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None
        self._timer_start = time.perf_counter()
        self.timer_var.set("Elapsed: 00:00:00")
        self._schedule_timer_tick()

    def _schedule_timer_tick(self) -> None:
        if self._timer_start is None:
            return
        elapsed = time.perf_counter() - self._timer_start
        self.timer_var.set(f"Elapsed: {self._format_elapsed(elapsed)}")
        self._timer_job = self.after(1000, self._schedule_timer_tick)

    def _stop_timer(self) -> None:
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None
        if self._timer_start is not None:
            elapsed = time.perf_counter() - self._timer_start
            self.timer_var.set(f"Elapsed: {self._format_elapsed(elapsed)}")
            self._timer_start = None

    def _reset_timer_display(self) -> None:
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None
        self._timer_start = None
        self.timer_var.set("Elapsed: 00:00:00")


def launch_gui() -> None:
    app = CerberusApp()
    app.mainloop()
