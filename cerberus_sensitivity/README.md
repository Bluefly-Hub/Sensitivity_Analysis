# Cerberus Sensitivity Python Runner

This package mirrors the PAD Cerberus Sensitivity automation using Python + UIAutomation and Tkinter.

## Modules

- engine.py: High-level entry point that wires the window hygiene, UI handle cache, trigger agent, and orchestrator.
- utomation/: Core agents responsible for window management, control caching, triggering UI actions, parsing clipboard tables, and building run plans.
- gui/app.py: Tkinter front-end that collects inputs (manual entry or paste from Excel), manages a background worker, updates progress, and allows copying the results as TSV.

## Usage

`ash
python -m cerberus_sensitivity.main
`

Or import in code:

`python
from cerberus_sensitivity import CerberusEngine
engine = CerberusEngine()
engine.run_scan(progress, data_list, start_index, cancel_event)
`

The progress argument should be a ProgressReporter instance. It emits init, 
ow, and done events that the GUI consumes.
