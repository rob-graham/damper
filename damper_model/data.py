"""CSV-backed configuration loader for the damper solver."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple


@dataclass
class ElementDefinition:
    element_id: str
    type: str
    cd: float
    area_mm2: float


@dataclass
class TopologyRow:
    topology: str
    element_id: str
    from_node: str
    to_node: str
    direction: str
    type_override: Optional[str] = None
    cd_override: Optional[float] = None
    area_override_mm2: Optional[float] = None


class DataRepository:
    """Loads damper configuration data from CSV files."""

    def __init__(self, root: Path | str) -> None:
        self.root = Path(root)
        if not self.root.exists():
            raise FileNotFoundError(self.root)

        self.scalars = self._load_scalar_table(self.root / "scalars.csv")
        self.nodes = self._load_node_pressures(self.root / "nodes.csv")
        self.elements = self._load_elements(self.root / "elements.csv")
        self.topology_rows = self._load_topology(self.root / "topology.csv")
        self.viscosity_table = self._load_viscosity(self.root / "viscosity.csv")
        self.click_areas = self._load_click_areas(self.root / "click_area.csv")
        self.shim_stacks = self._load_shim_stacks(self.root / "shim_stacks.csv")

        settings_path = self.root / "solver_settings.csv"
        if settings_path.exists():
            self.solver_settings = self._load_scalar_table(settings_path)
        else:
            self.solver_settings = {
                "MinVelocity": 0.0,
                "MaxVelocity": 0.6,
                "Step": 0.1,
                "Direction": "Compression",
                "Topology": "Fork_OpenCartridge",
            }

    # ------------------------------------------------------------------
    # Scalar helpers
    def get_scalar_float(self, name: str, default: float | None = None) -> float:
        value = self.scalars.get(name, default)
        if value is None or value == "":
            raise KeyError(f"Scalar '{name}' is missing")
        if isinstance(value, (int, float)):
            return float(value)
        return float(value)

    def get_scalar_text(self, name: str, default: str | None = None) -> str:
        value = self.scalars.get(name, default)
        if value is None or value == "":
            raise KeyError(f"Scalar '{name}' is missing")
        return str(value)

    def get_solver_setting(self, name: str, default: float | str | None = None):
        value = self.solver_settings.get(name, default)
        if isinstance(value, str) and value == "":
            return default
        return value

    def get_solver_setting_float(self, name: str, default: float | None = None) -> float:
        value = self.get_solver_setting(name, default)
        if value is None:
            raise KeyError(f"Solver setting '{name}' is missing")
        if isinstance(value, (int, float)):
            return float(value)
        return float(value)

    # ------------------------------------------------------------------
    # Lookups
    def lookup_bleed_area(self, click_setting: float) -> Optional[float]:
        click = int(round(click_setting))
        return self.click_areas.get(click)

    def lookup_shim_stack_rate(self, code: str) -> float:
        entry = self.shim_stacks.get(code)
        if not entry:
            return 50.0
        return float(entry.get("stack_rate_n_per_mm", 50.0))

    def interpolate_viscosity(self, temperature_c: float) -> float:
        table = self.viscosity_table
        if not table:
            return 100.0
        if temperature_c <= table[0][0]:
            return table[0][1]
        for idx in range(1, len(table)):
            temp, value = table[idx]
            prev_temp, prev_value = table[idx - 1]
            if temperature_c <= temp:
                span = temp - prev_temp
                if span <= 0.0:
                    return value
                fraction = (temperature_c - prev_temp) / span
                return prev_value + fraction * (value - prev_value)
        return table[-1][1]

    # ------------------------------------------------------------------
    # Loaders
    def _read_csv_dicts(self, path: Path) -> List[Dict[str, str]]:
        if not path.exists():
            raise FileNotFoundError(path)
        with path.open(newline="", encoding="utf-8") as fh:
            reader = csv.DictReader(fh)
            rows: List[Dict[str, str]] = []
            for raw_row in reader:
                if not raw_row:
                    continue
                cleaned = {
                    (key or "").strip(): (value.strip() if value is not None else "")
                    for key, value in raw_row.items()
                    if key
                }
                if any(cleaned.values()):
                    rows.append(cleaned)
            return rows

    def _parse_value(self, text: str) -> float | str | bool | int | float:
        stripped = text.strip()
        if stripped == "":
            return ""
        lowered = stripped.lower()
        if lowered in {"true", "false"}:
            return lowered == "true"
        try:
            return float(stripped)
        except ValueError:
            return stripped

    def _maybe_float(self, text: str | None) -> Optional[float]:
        if text is None:
            return None
        stripped = text.strip()
        if stripped == "":
            return None
        return float(stripped)

    def _load_scalar_table(self, path: Path) -> Dict[str, float | str | bool]:
        table: Dict[str, float | str | bool] = {}
        for row in self._read_csv_dicts(path):
            key = row.get("name") or row.get("key")
            if not key:
                continue
            table[key] = self._parse_value(row.get("value", ""))
        return table

    def _load_node_pressures(self, path: Path) -> Dict[str, float]:
        nodes: Dict[str, float] = {}
        for row in self._read_csv_dicts(path):
            node = row.get("node")
            pressure = row.get("pressure_bar") or row.get("pressure")
            if not node or not pressure:
                continue
            nodes[node] = float(pressure)
        return nodes

    def _load_elements(self, path: Path) -> Dict[str, ElementDefinition]:
        elements: Dict[str, ElementDefinition] = {}
        for row in self._read_csv_dicts(path):
            element_id = row.get("element_id")
            if not element_id:
                continue
            elements[element_id] = ElementDefinition(
                element_id=element_id,
                type=row.get("type", ""),
                cd=float(row.get("cd", "0") or 0.0),
                area_mm2=float(row.get("area_mm2", "0") or 0.0),
            )
        return elements

    def _load_topology(self, path: Path) -> List[TopologyRow]:
        rows: List[TopologyRow] = []
        for row in self._read_csv_dicts(path):
            element_id = row.get("element_id")
            from_node = row.get("from_node")
            to_node = row.get("to_node")
            if not element_id or not from_node or not to_node:
                continue
            rows.append(
                TopologyRow(
                    topology=row.get("topology", "*") or "*",
                    element_id=element_id,
                    from_node=from_node,
                    to_node=to_node,
                    direction=row.get("direction", ""),
                    type_override=row.get("type_override") or None,
                    cd_override=self._maybe_float(row.get("cd_override")),
                    area_override_mm2=self._maybe_float(row.get("area_override_mm2")),
                )
            )
        return rows

    def _load_viscosity(self, path: Path) -> List[Tuple[float, float]]:
        entries: List[Tuple[float, float]] = []
        for row in self._read_csv_dicts(path):
            temp = row.get("temp_c") or row.get("temperature")
            visc = row.get("viscosity_cst") or row.get("viscosity")
            if not temp or not visc:
                continue
            entries.append((float(temp), float(visc)))
        entries.sort(key=lambda item: item[0])
        return entries

    def _load_click_areas(self, path: Path) -> Dict[int, float]:
        areas: Dict[int, float] = {}
        for row in self._read_csv_dicts(path):
            click = row.get("click")
            area = row.get("bleed_area_mm2") or row.get("area_mm2")
            if not click or not area:
                continue
            areas[int(float(click))] = float(area)
        return areas

    def _load_shim_stacks(self, path: Path) -> Dict[str, Dict[str, float | str]]:
        stacks: Dict[str, Dict[str, float | str]] = {}
        for row in self._read_csv_dicts(path):
            code = row.get("code")
            if not code:
                continue
            stacks[code] = {
                "description": row.get("description", ""),
                "stack_rate_n_per_mm": float(row.get("stack_rate_n_per_mm", "50") or 50.0),
            }
        return stacks


__all__ = [
    "DataRepository",
    "ElementDefinition",
    "TopologyRow",
]
