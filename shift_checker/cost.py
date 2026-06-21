from __future__ import annotations

import datetime
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

from shift_checker.config import (
    build_employee_lookup,
    load_cost_rates,
    load_employees,
    load_holidays,
)
from shift_checker.parser import ShiftRecord


@dataclass
class ShiftCost:
    employee_id: str
    employee_name: str
    store: str
    position: str
    date: datetime.date
    hours: float
    hourly_rate: float
    day_type: str
    multiplier: float
    base_cost: float
    overtime_hours: float
    overtime_cost: float
    total_cost: float


@dataclass
class EmployeeCostSummary:
    employee_id: str
    employee_name: str
    store: str
    position: str
    total_hours: float
    regular_hours: float
    overtime_hours: float
    regular_cost: float
    overtime_cost: float
    total_cost: float
    by_day_type: dict[str, dict] = field(default_factory=dict)


@dataclass
class StoreCostSummary:
    store: str
    total_hours: float
    regular_cost: float
    overtime_cost: float
    total_cost: float
    employee_count: int


def _day_type(
    d: datetime.date,
    holidays: dict[str, str],
) -> str:
    ds = d.strftime("%Y-%m-%d")
    if ds in holidays:
        return "holiday"
    if d.weekday() >= 5:
        return "weekend"
    return "weekday"


def _get_multiplier(day_type: str, rates: dict) -> float:
    if day_type == "holiday":
        return rates.get("holiday_multiplier", 3.0)
    if day_type == "weekend":
        return rates.get("weekend_multiplier", 1.5)
    return rates.get("weekday_multiplier", 1.0)


def _get_overtime_multiplier(day_type: str, rates: dict) -> float:
    if day_type == "holiday":
        return rates.get("holiday_overtime_multiplier", 3.0)
    if day_type == "weekend":
        return rates.get("weekend_overtime_multiplier", 2.0)
    return rates.get("overtime_multiplier", 1.5)


def compute_shift_costs(
    records: list[ShiftRecord],
    employees: list[dict] | None = None,
    holidays: dict[str, str] | None = None,
    rates: dict | None = None,
) -> list[ShiftCost]:
    if employees is None:
        employees = load_employees()
    if holidays is None:
        holidays = load_holidays()
    if rates is None:
        rates = load_cost_rates()

    emp_lookup = build_employee_lookup(employees)
    standard_daily = rates.get("standard_daily_hours", 8.0)

    daily_hours: dict[tuple[str, datetime.date], float] = defaultdict(float)
    for r in records:
        daily_hours[(r.employee_id, r.date)] += r.hours

    results: list[ShiftCost] = []
    for r in records:
        emp = emp_lookup.get("by_id", {}).get(r.employee_id)
        hourly_rate = emp.get("hourly_rate", 0) if emp else 0
        position = emp.get("position", "未知") if emp else "未知"

        dt = _day_type(r.date, holidays)
        multiplier = _get_multiplier(dt, rates)
        base_cost = r.hours * hourly_rate * multiplier

        total_daily = daily_hours[(r.employee_id, r.date)]
        overtime_hours = max(0, total_daily - standard_daily)
        proportion = r.hours / total_daily if total_daily > 0 else 0
        overtime_h = overtime_hours * proportion

        ot_multiplier = _get_overtime_multiplier(dt, rates)
        overtime_cost = overtime_h * hourly_rate * (ot_multiplier - multiplier)

        total_cost = base_cost + overtime_cost

        results.append(ShiftCost(
            employee_id=r.employee_id,
            employee_name=r.employee_name,
            store=r.store,
            position=position,
            date=r.date,
            hours=r.hours,
            hourly_rate=hourly_rate,
            day_type=dt,
            multiplier=multiplier,
            base_cost=base_cost,
            overtime_hours=overtime_h,
            overtime_cost=overtime_cost,
            total_cost=total_cost,
        ))

    return results


def summarize_by_employee(costs: list[ShiftCost]) -> list[EmployeeCostSummary]:
    by_emp: dict[str, list[ShiftCost]] = defaultdict(list)
    for c in costs:
        by_emp[c.employee_id].append(c)

    summaries: list[EmployeeCostSummary] = []
    for eid, items in by_emp.items():
        total_hours = sum(i.hours for i in items)
        regular_hours = total_hours - sum(i.overtime_hours for i in items)
        overtime_hours = sum(i.overtime_hours for i in items)
        regular_cost = sum(i.base_cost for i in items)
        overtime_cost = sum(i.overtime_cost for i in items)
        total_cost = sum(i.total_cost for i in items)

        by_dt: dict[str, dict] = defaultdict(lambda: {"hours": 0, "cost": 0})
        for i in items:
            by_dt[i.day_type]["hours"] += i.hours
            by_dt[i.day_type]["cost"] += i.total_cost

        summaries.append(EmployeeCostSummary(
            employee_id=eid,
            employee_name=items[0].employee_name,
            store=items[0].store,
            position=items[0].position,
            total_hours=total_hours,
            regular_hours=regular_hours,
            overtime_hours=overtime_hours,
            regular_cost=regular_cost,
            overtime_cost=overtime_cost,
            total_cost=total_cost,
            by_day_type=dict(by_dt),
        ))

    return sorted(summaries, key=lambda s: s.employee_id)


def summarize_by_store(costs: list[ShiftCost]) -> list[StoreCostSummary]:
    by_store: dict[str, list[ShiftCost]] = defaultdict(list)
    for c in costs:
        by_store[c.store].append(c)

    summaries: list[StoreCostSummary] = []
    for store, items in by_store.items():
        emps = set(i.employee_id for i in items)
        summaries.append(StoreCostSummary(
            store=store,
            total_hours=sum(i.hours for i in items),
            regular_cost=sum(i.base_cost for i in items),
            overtime_cost=sum(i.overtime_cost for i in items),
            total_cost=sum(i.total_cost for i in items),
            employee_count=len(emps),
        ))

    return sorted(summaries, key=lambda s: s.store)
