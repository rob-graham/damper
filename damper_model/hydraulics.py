"""Hydraulic helper formulas used by the damper solver."""

from __future__ import annotations

import math

BAR_TO_PA = 1.0e5
MM2_TO_M2 = 1.0e-6
CM2_TO_M2 = 1.0e-4
MIN_CAVITATION_BAR = 0.3


def dp_orifice(flow_m3s: float, rho: float, cd: float, area_mm2: float) -> float:
    """Return the pressure drop across an orifice in bar."""
    area_m2 = area_mm2 * MM2_TO_M2
    if area_m2 <= 0.0 or rho <= 0.0 or cd <= 0.0:
        return 0.0
    velocity = flow_m3s / (cd * area_m2)
    dp_pa = 0.5 * rho * velocity * velocity
    return dp_pa / BAR_TO_PA


def flow_orifice(delta_p_bar: float, rho: float, cd: float, area_mm2: float) -> float:
    """Return the volumetric flow through an orifice for a pressure drop."""
    area_m2 = area_mm2 * MM2_TO_M2
    if area_m2 <= 0.0 or rho <= 0.0 or cd <= 0.0:
        return 0.0
    delta_p_pa = delta_p_bar * BAR_TO_PA
    if delta_p_pa == 0.0:
        return 0.0
    sign = 1.0 if delta_p_pa > 0.0 else -1.0
    return sign * cd * area_m2 * math.sqrt(2.0 * abs(delta_p_pa) / rho)


def dp_poiseuille(flow_m3s: float, mu_pas: float, length_m: float, diameter_m: float) -> float:
    """Return the laminar pressure drop of a capillary in bar."""
    if mu_pas <= 0.0 or length_m <= 0.0 or diameter_m <= 0.0:
        return 0.0
    numerator = 128.0 * mu_pas * length_m * flow_m3s
    denominator = math.pi * diameter_m**4
    dp_pa = numerator / denominator
    return dp_pa / BAR_TO_PA


def dp_blend(
    flow_m3s: float,
    rho: float,
    mu_pas: float,
    area_mm2: float,
    cd: float,
    length_m: float | None = None,
) -> float:
    """Blend laminar and turbulent pressure drops for a small passage."""
    area_m2 = area_mm2 * MM2_TO_M2
    if area_m2 <= 0.0 or rho <= 0.0 or cd <= 0.0:
        return 0.0

    diameter_m = math.sqrt(4.0 * area_m2 / math.pi)
    effective_length = length_m if length_m and length_m > 0.0 else 3.0 * diameter_m
    velocity = flow_m3s / area_m2 if area_m2 > 0.0 else 0.0

    reynolds = 0.0
    if mu_pas > 0.0:
        reynolds = rho * abs(velocity) * diameter_m / mu_pas

    laminar_dp = dp_poiseuille(flow_m3s, mu_pas, effective_length, diameter_m)
    orifice_dp = dp_orifice(flow_m3s, rho, cd, area_mm2)

    weight = max(0.0, min(1.0, (reynolds - 1500.0) / (3000.0 - 1500.0)))
    return (1.0 - weight) * laminar_dp + weight * orifice_dp


def shim_lift(
    delta_p_bar: float,
    area_mm2: float,
    stack_rate_n_per_mm: float,
    preload_n: float = 0.0,
) -> float:
    """Return shim deflection in millimetres for a given pressure drop."""
    if stack_rate_n_per_mm <= 0.0:
        return 0.0
    area_m2 = area_mm2 * MM2_TO_M2
    force_n = delta_p_bar * BAR_TO_PA * area_m2
    if force_n <= preload_n:
        return 0.0
    return (force_n - preload_n) / stack_rate_n_per_mm


__all__ = [
    "BAR_TO_PA",
    "MM2_TO_M2",
    "CM2_TO_M2",
    "MIN_CAVITATION_BAR",
    "dp_orifice",
    "flow_orifice",
    "dp_poiseuille",
    "dp_blend",
    "shim_lift",
]
