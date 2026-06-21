import json
from pathlib import Path
from typing import Any

CONFIG_DIR = Path(__file__).parent / "config"


def _load_json(path: Path) -> Any:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_column_synonyms(path: Path | None = None) -> dict[str, list[str]]:
    p = path or CONFIG_DIR / "column_synonyms.json"
    return _load_json(p)


def load_shift_codes(path: Path | None = None) -> dict[str, dict]:
    p = path or CONFIG_DIR / "shift_codes.json"
    return _load_json(p)


def load_employees(path: Path | None = None) -> list[dict]:
    p = path or CONFIG_DIR / "employees.json"
    return _load_json(p)


def load_holidays(path: Path | None = None) -> dict[str, str]:
    p = path or CONFIG_DIR / "holidays.json"
    return _load_json(p)


def load_rules(path: Path | None = None) -> dict:
    p = path or CONFIG_DIR / "rules.json"
    return _load_json(p)


def load_cost_rates(path: Path | None = None) -> dict:
    p = path or CONFIG_DIR / "cost_rates.json"
    return _load_json(p)


def build_employee_lookup(employees: list[dict]) -> dict:
    by_id: dict[str, dict] = {}
    alias_to_id: dict[str, str] = {}
    for emp in employees:
        eid = emp["employee_id"]
        by_id[eid] = emp
        alias_to_id[eid] = eid
        for alias in emp.get("aliases", []):
            alias_to_id[alias] = eid
        alias_to_id[emp["employee_name"]] = eid
    return {"by_id": by_id, "alias_to_id": alias_to_id}


def build_column_lookup(synonyms: dict[str, list[str]]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for canonical, syns in synonyms.items():
        lookup[canonical] = canonical
        for s in syns:
            lookup[s] = canonical
    return lookup
