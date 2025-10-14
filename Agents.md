# Agents.md — Python Rebuild of PAD Automations

This file outlines how to rebuild **PAD flows** like *Cerberus Sensitivity* into Python, based on the same structure you used for *Liner Hanger Flow*. It’s not meant to be exact — instead it sketches the moving pieces (agents) and how they should interact.

---

## 1) Overall Pattern

* **Engine/Automation Layer**: Handles window attach, control discovery, setting values, triggering calculations, reading outputs.
* **GUI Layer**: Tkinter app that collects inputs, shows progress, manages threading, handles errors, and exports results.
* **Progress Bus**: Engine emits `init/row/done` events; GUI consumes them.

This matches your working Python code: `Liner_Hanger_Flow.py` (engine) + `Wellplan_gui.py` (GUI).

---

## 2) Core Agents

### A. Orchestrator Agent

* Accepts inputs from GUI (lists of numbers, parameters).
* Runs the scan loop: set value → retrigger calc → wait for output → collect.
* Emits structured progress messages to the GUI.
* Handles resume by starting at `start_index`.

### B. UI Handles Agent

* Caches references to key controls (edits, radios, checkboxes, tabs).
* Provides safe getters with lazy refresh.
* Expands/scrolls into view as needed.
* Only button_repository.y can be used for the API.

### C. UI Trigger Agent

* Sets values via `ValuePattern` (no keystrokes).
* Uses `SelectionItemPattern` for toggling radios/checkboxes.
* Retriggers calculations if outputs don’t change.
* Waits for output changes with polling + timeout.

### D. Window Hygiene Agent

* Attaches to the main app window via regex title match.
* Brings window forward/restores when needed.
* Scrolls controls into view.

### E. Error/Recovery Agent

* Cooperative cancel via threading event.
* Duplicate detection and resume cleanup.
* Error dialogs with Resume/Stop options.

### F. Export Agent

* Collects results into a simple table.
* Provides a **Copy** button that puts TSV on the clipboard.

---

## 3) GUI Responsibilities

* **Inputs**: Treeview with editable rows or pasted structured values (you’ll replace clipboard parsing with GUI inputs for Cerberus).
* **Run/Resume/Stop**: Buttons wired to worker thread.
* **Results**: Table view showing (index, inputs, output).
* **Status**: Updates based on progress events.
* **Copy**: Places results on clipboard.
* **Error Dialogs**: Friendly recovery options.

---

## 4) Cerberus Sensitivity Flow (PAD → Python)

* Open **Sensitivity Analysis** menu.
* Load template (e.g., *auto*).
* Toggle tabs/checkboxes (RIH, POOH, Pipe fluid density, Depth).
* Open **Parameter Matrix Wizard**:

  * Clear values, add Depth slice.
  * On first run, also add Pipe Fluid Density and FOE.
* Confirm, Calculate.
* Collect results (tabulated data in PAD → in Python, scrape via UIA or read outputs directly).
* If total iterations > 200, batch depth values into multiple runs.
* GUI replaces PAD’s Excel macro: display results, copy to clipboard.

---

## 5) Flow of Data

```
GUI → Orchestrator (inputs) → UI Handles → UI Trigger → Target App
GUI ← Progress Bus ← Orchestrator (outputs)
GUI → Export Agent → Clipboard
```

---

## 6) Error Handling Principles

* Prefer UIA patterns over clicks.
* Skip setting values if already equal.
* Use short polling with hard timeouts.
* Retrigger fallback (toggle Top-down → Bottom-up, or RIH → POOH).
* Foreground/scroll hygiene before every action.
* Expose clear errors to GUI, allow resume.

---

## 7) What to Keep Consistent

* Engine/GUI separation.
* `run_scan(progress, data_list, start_index, cancel_event)` signature.
* Standard progress events: `init`, `row`, `done`.
* Tkinter pattern: Treeviews, buttons, status label, worker thread.
* Keep-awake background thread for long runs.

---

## 8) Porting Checklist

1. Define stable window match string.
2. Identify AutomationIds/Names for controls.
3. Implement UI Handles wrapper.
4. Build Orchestrator loop with retries/timeouts.
5. Design GUI Treeview for inputs.
6. Map PAD batching logic (e.g., Depth > 200 values → chunk runs).
7. Provide TSV Copy button.

---

## 9) Example Minimal Interface

```python
final_inputs, outputs = run_scan(
    progress=my_progress,
    data_list=[...],
    start_index=0,
    cancel_event=my_event
)
```

---

This is the blueprint for remaking *Cerberus Sensitivity* and similar PAD flows: keep your current Python structure, but shift inputs to GUI instead of clipboard, and adapt the loop to match each PAD sequence.
