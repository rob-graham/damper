"""Damper network solver."""

from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List

from .data import DataRepository, ElementDefinition, TopologyRow
from .hydraulics import (
    BAR_TO_PA,
    CM2_TO_M2,
    MIN_CAVITATION_BAR,
    flow_orifice,
    dp_blend,
    shim_lift,
)

MAX_ITER = 50
TOL_RESIDUAL = 1.0e-9
FD_STEP = 1.0e-5


@dataclass
class Branch:
    element_id: str
    type: str
    cd: float
    area_mm2: float
    from_node: str
    to_node: str


@dataclass
class DamperEvaluation:
    residual: List[float]
    pressures: Dict[str, float]
    flows: Dict[str, float]
    net_flow: Dict[str, float]


@dataclass
class DamperResult:
    node_pressures: Dict[str, float]
    element_flows: Dict[str, float]
    force_n: float
    delta_p_bar: float
    losses_bar: float
    cavitation_margin_bar: float
    bleed_fraction: float
    shim_lift_mm: float

    def as_row(self, velocity: float) -> Dict[str, float]:
        return {
            "velocity_m_per_s": velocity,
            "force_N": self.force_n,
            "deltaP_bar": self.delta_p_bar,
            "losses_bar": self.losses_bar,
            "cavitation_margin_bar": self.cavitation_margin_bar,
            "bleed_fraction": self.bleed_fraction,
            "shim_lift_mm": self.shim_lift_mm,
        }


class DamperSolver:
    def __init__(self, data: DataRepository) -> None:
        self.data = data

    # ------------------------------------------------------------------
    def solve(
        self,
        direction: str,
        velocity_m_per_s: float,
        travel_position: float,
        topology_name: str,
    ) -> DamperResult:
        topology = self._build_configuration(direction, topology_name)
        nodes = dict(self.data.nodes)

        rho = self.data.get_scalar_float("OilDensity")
        viscosity_cst = self.data.interpolate_viscosity(self.data.get_scalar_float("Temperature"))
        mu_pas = viscosity_cst * 1.0e-6 * rho
        ap_m2 = self.data.get_scalar_float("Ap") * CM2_TO_M2
        ar_m2 = self.data.get_scalar_float("Ar") * CM2_TO_M2
        dir_sign = -1.0 if direction.lower() == "rebound" else 1.0

        fixed_nodes = {
            "Shaft": self.data.get_scalar_float("ShaftPressure"),
            "Reservoir": self.data.get_scalar_float("ReservoirPressure"),
        }

        unknown_names = [name for name in nodes.keys() if name not in fixed_nodes]
        unknown_values = [nodes[name] for name in unknown_names]

        if unknown_names:
            values = unknown_values[:]
            converged = False
            for _ in range(MAX_ITER):
                evaluation = self._evaluate_network(
                    unknown_names,
                    values,
                    nodes,
                    fixed_nodes,
                    topology,
                    rho,
                    mu_pas,
                    dir_sign,
                    velocity_m_per_s,
                    ap_m2,
                    ar_m2,
                )
                residual = evaluation.residual
                max_residual = max((abs(value) for value in residual), default=0.0)
                if max_residual < TOL_RESIDUAL:
                    converged = True
                    break
                jacobian = self._numerical_jacobian(
                    unknown_names,
                    values,
                    nodes,
                    fixed_nodes,
                    topology,
                    rho,
                    mu_pas,
                    dir_sign,
                    velocity_m_per_s,
                    ap_m2,
                    ar_m2,
                )
                delta = self._solve_linear_system(jacobian, residual)
                for index in range(len(values)):
                    values[index] -= delta[index]
                    if values[index] < MIN_CAVITATION_BAR:
                        values[index] = MIN_CAVITATION_BAR
            if not converged:
                evaluation = self._evaluate_network(
                    unknown_names,
                    values,
                    nodes,
                    fixed_nodes,
                    topology,
                    rho,
                    mu_pas,
                    dir_sign,
                    velocity_m_per_s,
                    ap_m2,
                    ar_m2,
                )
        else:
            evaluation = self._evaluate_network(
                unknown_names,
                [],
                nodes,
                fixed_nodes,
                topology,
                rho,
                mu_pas,
                dir_sign,
                velocity_m_per_s,
                ap_m2,
                ar_m2,
            )

        pressures = evaluation.pressures
        flows = evaluation.flows

        p_a = pressures.get("ChamberA", 0.0)
        p_b = pressures.get("ChamberB", 0.0)
        p_r = pressures.get("Reservoir", 0.0)

        force_n = (p_a - p_b) * BAR_TO_PA * ap_m2 + (p_b - p_r) * BAR_TO_PA * ar_m2
        delta_p = p_a - p_b
        losses = self._compute_losses(topology, pressures)
        cav_margin = min(pressures.values()) - MIN_CAVITATION_BAR if pressures else 0.0
        bleed_fraction = self._compute_bleed_fraction(topology, flows)

        shim_stack_code = self.data.get_scalar_text("ShimStack")
        stack_rate = self.data.lookup_shim_stack_rate(shim_stack_code)
        compression_branch = self._get_branch_by_type(topology, "Compression")
        shim_lift_mm = 0.0
        if compression_branch:
            comp_delta_p = pressures[compression_branch.from_node] - pressures[compression_branch.to_node]
            shim_lift_mm = shim_lift(comp_delta_p, compression_branch.area_mm2, stack_rate)

        return DamperResult(
            node_pressures=pressures,
            element_flows=flows,
            force_n=force_n,
            delta_p_bar=delta_p,
            losses_bar=losses,
            cavitation_margin_bar=cav_margin,
            bleed_fraction=bleed_fraction,
            shim_lift_mm=shim_lift_mm,
        )

    # ------------------------------------------------------------------
    def run_sweep(
        self,
        direction: str | None = None,
        topology: str | None = None,
        velocity_min: float | None = None,
        velocity_max: float | None = None,
        step: float | None = None,
        travel_position: float | None = None,
    ) -> List[Dict[str, float]]:
        direction = direction or str(self.data.get_solver_setting("Direction", "Compression"))
        topology = topology or str(self.data.get_solver_setting("Topology", "Fork_OpenCartridge"))
        v_min = velocity_min if velocity_min is not None else self.data.get_solver_setting_float("MinVelocity", 0.0)
        v_max = velocity_max if velocity_max is not None else self.data.get_solver_setting_float("MaxVelocity", 0.6)
        step = step if step is not None else self.data.get_solver_setting_float("Step", 0.1)
        if step <= 0.0:
            raise ValueError("Velocity step must be positive")

        travel_position = (
            travel_position
            if travel_position is not None
            else self.data.get_scalar_float("TravelPosition")
        )

        velocities: List[float] = []
        if v_max >= v_min:
            value = v_min
            while value <= v_max + 1e-9:
                velocities.append(value)
                value += step
        else:
            value = v_min
            while value >= v_max - 1e-9:
                velocities.append(value)
                value -= step

        rows: List[Dict[str, float]] = []
        for velocity in velocities:
            result = self.solve(direction, abs(velocity), travel_position, topology)
            row = result.as_row(velocity)
            rows.append(row)
        return rows

    # ------------------------------------------------------------------
    def _build_configuration(self, direction: str, topology_name: str) -> Dict[str, Branch]:
        elements = self.data.elements
        topology: Dict[str, Branch] = {}

        for row in self.data.topology_rows:
            if row.topology not in ("*", "", topology_name):
                continue
            if not self._should_include_branch(row.direction, direction):
                continue

            base_element: ElementDefinition | None = elements.get(row.element_id)
            if base_element is None and not (row.type_override and row.cd_override and row.area_override_mm2):
                raise KeyError(f"Element '{row.element_id}' is missing from elements.csv")

            branch_type = row.type_override or (base_element.type if base_element else "")
            cd = row.cd_override if row.cd_override is not None else (base_element.cd if base_element else 0.0)
            area = (
                row.area_override_mm2
                if row.area_override_mm2 is not None
                else (base_element.area_mm2 if base_element else 0.0)
            )

            topology[row.element_id] = Branch(
                element_id=row.element_id,
                type=branch_type,
                cd=cd,
                area_mm2=area,
                from_node=row.from_node,
                to_node=row.to_node,
            )

        self._apply_extensions(topology, topology_name, direction)
        if not topology:
            raise ValueError(f"No active branches for topology '{topology_name}' in direction '{direction}'")
        return topology

    def _apply_extensions(self, topology: Dict[str, Branch], topology_name: str, direction: str) -> None:
        bleed_area = self.data.lookup_bleed_area(self.data.get_scalar_float("ClickSetting"))
        if bleed_area is not None:
            for branch in topology.values():
                if branch.type.lower() == "bleed":
                    branch.area_mm2 = bleed_area

        if topology_name.lower() == "shock_remoteres" and "RR1" not in topology:
            cd = float(self.data.get_scalar_float("RemoteCd", 0.55))
            area = float(self.data.get_scalar_float("RemoteArea_mm2", 0.45))
            if direction.lower() == "compression":
                from_node, to_node = "ChamberA", "Reservoir"
            else:
                from_node, to_node = "Reservoir", "ChamberA"
            topology["RR1"] = Branch(
                element_id="RR1",
                type="Remote",
                cd=cd,
                area_mm2=area,
                from_node=from_node,
                to_node=to_node,
            )

    def _should_include_branch(self, branch_direction: str, request_direction: str) -> bool:
        branch_upper = branch_direction.strip().upper()
        if branch_upper in {"BIDIRECTIONAL", "BOTH", "COMMON", "BLEED"}:
            return True
        return branch_direction.strip().lower() == request_direction.strip().lower()

    # ------------------------------------------------------------------
    def _evaluate_network(
        self,
        unknown_names: List[str],
        unknown_values: List[float],
        nodes: Dict[str, float],
        fixed_nodes: Dict[str, float],
        topology: Dict[str, Branch],
        rho: float,
        mu_pas: float,
        dir_sign: float,
        velocity: float,
        ap_m2: float,
        ar_m2: float,
    ) -> DamperEvaluation:
        pressures: Dict[str, float] = dict(nodes)
        for name, value in zip(unknown_names, unknown_values):
            pressures[name] = value
        for name, value in fixed_nodes.items():
            pressures[name] = value

        flows: Dict[str, float] = {}
        net_flow: Dict[str, float] = defaultdict(float)

        for branch in topology.values():
            delta_p = pressures[branch.from_node] - pressures[branch.to_node]
            flow = self._compute_element_flow(branch, delta_p, rho, mu_pas)
            flows[branch.element_id] = flow
            net_flow[branch.from_node] -= flow
            net_flow[branch.to_node] += flow

        q_a = -dir_sign * velocity * ap_m2
        q_b = dir_sign * velocity * (ap_m2 - ar_m2)
        net_flow["ChamberA"] += q_a
        net_flow["ChamberB"] += q_b

        residual = [net_flow[name] for name in unknown_names]
        return DamperEvaluation(
            residual=residual,
            pressures=dict(pressures),
            flows=flows,
            net_flow=dict(net_flow),
        )

    def _numerical_jacobian(
        self,
        unknown_names: List[str],
        unknown_values: List[float],
        nodes: Dict[str, float],
        fixed_nodes: Dict[str, float],
        topology: Dict[str, Branch],
        rho: float,
        mu_pas: float,
        dir_sign: float,
        velocity: float,
        ap_m2: float,
        ar_m2: float,
    ) -> List[List[float]]:
        n = len(unknown_values)
        jac = [[0.0 for _ in range(n)] for _ in range(n)]
        for j in range(n):
            perturbed = unknown_values[:]
            perturbed[j] += FD_STEP
            res_plus = self._evaluate_network(
                unknown_names,
                perturbed,
                nodes,
                fixed_nodes,
                topology,
                rho,
                mu_pas,
                dir_sign,
                velocity,
                ap_m2,
                ar_m2,
            ).residual

            perturbed[j] -= 2.0 * FD_STEP
            res_minus = self._evaluate_network(
                unknown_names,
                perturbed,
                nodes,
                fixed_nodes,
                topology,
                rho,
                mu_pas,
                dir_sign,
                velocity,
                ap_m2,
                ar_m2,
            ).residual

            for i in range(n):
                jac[i][j] = (res_plus[i] - res_minus[i]) / (2.0 * FD_STEP)
        return jac

    def _solve_linear_system(self, matrix: List[List[float]], rhs: List[float]) -> List[float]:
        n = len(rhs)
        a = [row[:] for row in matrix]
        b = rhs[:]
        for i in range(n):
            pivot = a[i][i]
            if abs(pivot) < 1.0e-12:
                pivot = 1.0e-12
            inv_pivot = 1.0 / pivot
            for j in range(i, n):
                a[i][j] *= inv_pivot
            b[i] *= inv_pivot
            for k in range(n):
                if k == i:
                    continue
                factor = a[k][i]
                if factor == 0.0:
                    continue
                for j in range(i, n):
                    a[k][j] -= factor * a[i][j]
                b[k] -= factor * b[i]
        return b

    def _compute_element_flow(self, branch: Branch, delta_p_bar: float, rho: float, mu_pas: float) -> float:
        element_type = branch.type.lower()
        if element_type in {"bleed", "remote"}:
            return self._solve_flow_from_dp(delta_p_bar, rho, mu_pas, branch.area_mm2, branch.cd)
        if element_type in {"capillary", "laminar"}:
            area_m2 = branch.area_mm2 * 1.0e-6
            if area_m2 <= 0.0:
                return 0.0
            diameter = math.sqrt(4.0 * area_m2 / math.pi)
            length = 3.0 * diameter
            flow = delta_p_bar * BAR_TO_PA
            flow /= 128.0 * mu_pas * length / (math.pi * diameter**4)
            return flow
        return flow_orifice(delta_p_bar, rho, branch.cd, branch.area_mm2)

    def _solve_flow_from_dp(
        self,
        delta_p_bar: float,
        rho: float,
        mu_pas: float,
        area_mm2: float,
        cd: float,
    ) -> float:
        if abs(delta_p_bar) < 1.0e-12:
            return 0.0
        target = delta_p_bar
        guess = 1.0e-6 if delta_p_bar >= 0.0 else -1.0e-6
        flow = guess
        for _ in range(40):
            f = dp_blend(flow, rho, mu_pas, area_mm2, cd) - target
            if abs(f) < 1.0e-9:
                break
            df = (
                dp_blend(flow + 1.0e-6, rho, mu_pas, area_mm2, cd)
                - dp_blend(flow - 1.0e-6, rho, mu_pas, area_mm2, cd)
            ) / (2.0e-6)
            if abs(df) < 1.0e-12:
                break
            flow -= f / df
        return flow

    def _compute_losses(self, topology: Dict[str, Branch], pressures: Dict[str, float]) -> float:
        total = 0.0
        for branch in topology.values():
            dp = pressures[branch.from_node] - pressures[branch.to_node]
            total += abs(dp)
        return max(0.0, total)

    def _compute_bleed_fraction(self, topology: Dict[str, Branch], flows: Dict[str, float]) -> float:
        total = 0.0
        bleed = 0.0
        for element_id, flow in flows.items():
            total += abs(flow)
            branch = topology.get(element_id)
            if branch and branch.type.lower() == "bleed":
                bleed += abs(flow)
        return bleed / total if total > 0.0 else 0.0

    def _get_branch_by_type(self, topology: Dict[str, Branch], branch_type: str) -> Branch | None:
        for branch in topology.values():
            if branch.type.lower() == branch_type.lower():
                return branch
        return None


__all__ = [
    "DamperSolver",
    "DamperResult",
]
