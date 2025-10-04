## Damper Solver Workbook

This repository contains a macro-enabled workflow for evaluating hydraulic damper configurations. The Excel workbook provides the input data, plots, and reporting surfaces while the VBA module implements the solver, pressure-drop correlations, and automation entry points.

## Repository layout

`damper_workbook.xlsx` – primary workbook containing inputs, plots, named ranges, and solver dashboards.
`vba/DamperSolver.bas` – exported VBA module with hydraulic utility functions (`DP_Orifice`, `DP_Blend`, `DP_Poiseuille`, `ShimLift`) plus the `SolveDamper`, `RunSweep`, `SinglePoint`, and `ExportCSV` routines.

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
+5. Use **Export CSV** to write the populated sweep table (`SolverResults`) to a comma-delimited file for further analysis.

The VBA solver enforces cavitation limits, supports both the `Fork_OpenCartridge` and `Shock_RemoteRes` topology configurations, and computes additional diagnostics such as bleed flow fractions and shim lift based on the active stack selection.
