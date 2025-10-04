"""Command-line interface for the damper solver."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import List

from .data import DataRepository
from .solver import DamperResult, DamperSolver


def main(argv: List[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Motorcycle damper solver")
    parser.add_argument("data_root", type=Path, help="Directory containing CSV configuration files")

    sub = parser.add_subparsers(dest="command", required=True)

    single = sub.add_parser("single", help="Run a single operating point")
    single.add_argument("--direction", choices=["Compression", "Rebound"])
    single.add_argument("--topology")
    single.add_argument("--velocity", type=float, help="Shaft velocity in m/s")
    single.add_argument("--travel", type=float, help="Stroke position ratio")
    single.add_argument("--output", type=Path, help="Optional JSON output path")

    sweep = sub.add_parser("sweep", help="Run a velocity sweep")
    sweep.add_argument("--direction", choices=["Compression", "Rebound"])
    sweep.add_argument("--topology")
    sweep.add_argument("--vmin", type=float, help="Minimum velocity")
    sweep.add_argument("--vmax", type=float, help="Maximum velocity")
    sweep.add_argument("--step", type=float, help="Sweep step size")
    sweep.add_argument("--travel", type=float, help="Stroke position ratio")
    sweep.add_argument("--output", type=Path, help="Optional CSV output path")

    args = parser.parse_args(argv)

    repo = DataRepository(args.data_root)
    solver = DamperSolver(repo)

    if args.command == "single":
        direction = args.direction or str(repo.get_solver_setting("Direction", "Compression"))
        topology = args.topology or str(repo.get_solver_setting("Topology", "Fork_OpenCartridge"))
        velocity = args.velocity if args.velocity is not None else float(repo.get_scalar_float("VelocityTarget"))
        travel = args.travel if args.travel is not None else repo.get_scalar_float("TravelPosition")
        result = solver.solve(direction, abs(velocity), travel, topology)
        _print_single_result(velocity, result)
        if args.output:
            _write_single_json(args.output, velocity, result)
    elif args.command == "sweep":
        rows = solver.run_sweep(
            direction=args.direction,
            topology=args.topology,
            velocity_min=args.vmin,
            velocity_max=args.vmax,
            step=args.step,
            travel_position=args.travel,
        )
        if not rows:
            print("Sweep produced no rows.")
            return
        _print_sweep_summary(rows)
        if args.output:
            _write_sweep_csv(args.output, rows)


def _print_single_result(velocity: float, result: DamperResult) -> None:
    print(f"Velocity [m/s]: {velocity:+.3f}")
    print(f"Force [N]: {result.force_n:.1f}")
    print(f"ΔP [bar]: {result.delta_p_bar:.3f}")
    print(f"Losses [bar]: {result.losses_bar:.3f}")
    print(f"Cavitation margin [bar]: {result.cavitation_margin_bar:.3f}")
    print(f"Bleed fraction [-]: {result.bleed_fraction:.3f}")
    print(f"Shim lift [mm]: {result.shim_lift_mm:.3f}")
    print("Node pressures [bar]:")
    for name, value in sorted(result.node_pressures.items()):
        print(f"  {name}: {value:.3f}")
    print("Element flows [m³/s]:")
    for name, value in sorted(result.element_flows.items()):
        print(f"  {name}: {value:.6e}")


def _write_single_json(path: Path, velocity: float, result: DamperResult) -> None:
    payload = {
        "velocity_m_per_s": velocity,
        "force_N": result.force_n,
        "deltaP_bar": result.delta_p_bar,
        "losses_bar": result.losses_bar,
        "cavitation_margin_bar": result.cavitation_margin_bar,
        "bleed_fraction": result.bleed_fraction,
        "shim_lift_mm": result.shim_lift_mm,
        "node_pressures_bar": result.node_pressures,
        "element_flows_m3s": result.element_flows,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, indent=2)
    print(f"Wrote single-point JSON to {path}")


def _print_sweep_summary(rows: List[dict]) -> None:
    print("Velocity sweep results:")
    for row in rows:
        print(
            f"  v={row['velocity_m_per_s']:+.3f} m/s "
            f"→ Force={row['force_N']:.1f} N, ΔP={row['deltaP_bar']:.3f} bar"
        )


def _write_sweep_csv(path: Path, rows: List[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote sweep CSV to {path}")


if __name__ == "__main__":  # pragma: no cover
    main()
