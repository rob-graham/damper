## Damper Solver Workbook

This repository contains a macro-enabled workflow for evaluating hydraulic damper configurations. The Excel workbook provides the input data, plots, and reporting surfaces while the VBA module implements the solver, pressure-drop correlations, and automation entry points.

## Repository layout

`damper_workbook.xlsx` – primary workbook containing inputs, plots, named ranges, and solver dashboards.
`vba/DamperSolver.bas` – exported VBA module with hydraulic utility functions (`DP_Orifice`, `DP_Blend`, `DP_Poiseuille`, `ShimLift`) plus the `SolveDamper`, `RunSweep`, `SinglePoint`, and `ExportCSV` routines.
`damper_model/` – pure-Python implementation of the hydraulic solver.
`data/` – CSV configuration bundle used by the Python solver for the default fork cartridge example.

## Installing the solver module

1. Open `damper_workbook.xlsx` in Excel.
2. Press `ALT + F11` to open the VBA editor. If a `DamperSolver` module already exists, remove or rename it.
3. From the VBA editor choose **File → Import File…** and select `vba/DamperSolver.bas` from this repository.
4. Save the workbook as a macro-enabled workbook (`.xlsm`) if prompted and enable macros when reopening.

Once the module is installed the Solver sheet buttons can be assigned to the public procedures:

**Run Sweep** → `DamperSolver.RunSweep`
**Single Point** → `DamperSolver.SinglePoint`
**Export CSV** → `DamperSolver.ExportCSV`

## Workflow overview

1. Populate the `Inputs`, `Elements`, `Topology`, and `Nodes` sheets. Named ranges used by the solver (e.g., `NodePressures`, `ElementDefinitions`, `TopologyMap`, `SolverSettings`) should remain intact.
2. Choose the active direction, topology, and sweep bounds in the `SolverSettings` block (cells `H2:I6`) on the Solver sheet.
3. Click **Run Sweep** to iterate over the configured velocity range. Charts on the Plots sheet will update automatically through the named ranges `ForceVelocity`, `DeltaPBudget`, and `FlowSplit` which now support up to 100 data rows.
4. Use **Single Point** to debug the current `VelocityTarget` from the Inputs sheet. The solver updates node pressures, flow splits, and summary metrics.
5. Use **Export CSV** to write the populated sweep table (`SolverResults`) to a comma-delimited file for further analysis.

The VBA solver enforces cavitation limits, supports both the `Fork_OpenCartridge` and `Shock_RemoteRes` topology configurations, and computes additional diagnostics such as bleed flow fractions and shim lift based on the active stack selection.

## Python solver

The `damper_model` package mirrors the workbook’s VBA routines so the same calculations can be run without Excel. The solver consumes CSV files that correspond to the workbook’s named ranges, making alternative configurations easy to version control.

### Running a single operating point

```bash
python -m damper_model.cli data single
```

### Running a sweep and exporting CSV

```bash
python -m damper_model.cli data sweep --direction Rebound --vmin -0.3 --vmax 0.0 --step 0.05 --output out/sweep.csv
```

Command-line overrides mirror the macro entry points: direction, topology, velocity bounds, and stroke position can all be provided explicitly; otherwise values from `data/scalars.csv` and `data/solver_settings.csv` are used.
