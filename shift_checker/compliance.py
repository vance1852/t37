from __future__ import annotations

import datetime
from collections import defaultdict
from dataclasses import dataclass, field
from itertools import groupby
from typing import Any

from shift_checker.config import build_employee_lookup, load_employees, load_rules
from shift_checker.parser import ShiftRecord


@dataclass
class Violation:
    employee_id: str
    employee_name: str
    rule: str
    details: str
    dates: list[datetime.date] = field(default_factory=list)
    severity: str = "warning"


def _shift_end_dt(rec: ShiftRecord) -> datetime.datetime:
    d = rec.date
    end = rec.end_time or datetime.time(5, 0)
    dt = datetime.datetime.combine(d, end)
    if rec.overnight:
        dt += datetime.timedelta(days=1)
    return dt


def _shift_start_dt(rec: ShiftRecord) -> datetime.datetime:
    return datetime.datetime.combine(rec.date, rec.start_time or datetime.time(9, 0))


def check_cross_store_overlap(records: list[ShiftRecord]) -> list[Violation]:
    violations: list[Violation] = []
    by_emp_date = defaultdict(list)
    for r in records:
        by_emp_date[(r.employee_id, r.date)].append(r)

    for (eid, d), recs in by_emp_date.items():
        if len(recs) <= 1:
            continue
        stores = set(r.store for r in recs)
        if len(stores) > 1:
            store_list = ", ".join(sorted(stores))
            violations.append(Violation(
                employee_id=eid,
                employee_name=recs[0].employee_name,
                rule="跨店重叠排班",
                details=f"同一天在多个门店排班: {store_list}",
                dates=[d],
                severity="critical",
            ))
        else:
            total = sum(r.hours for r in recs)
            if total > 12:
                violations.append(Violation(
                    employee_id=eid,
                    employee_name=recs[0].employee_name,
                    rule="同日重复排班",
                    details=f"同一天排班总时长 {total:.1f} 小时",
                    dates=[d],
                    severity="warning",
                ))
    return violations


def check_rest_insufficiency(
    records: list[ShiftRecord], min_rest_hours: float = 11.0,
) -> list[Violation]:
    violations: list[Violation] = []
    by_emp = defaultdict(list)
    for r in records:
        by_emp[r.employee_id].append(r)

    for eid, recs in by_emp.items():
        sorted_recs = sorted(recs, key=lambda x: _shift_start_dt(x))
        for i in range(len(sorted_recs) - 1):
            curr_end = _shift_end_dt(sorted_recs[i])
            next_start = _shift_start_dt(sorted_recs[i + 1])
            rest = (next_start - curr_end).total_seconds() / 3600
            if 0 < rest < min_rest_hours:
                violations.append(Violation(
                    employee_id=eid,
                    employee_name=sorted_recs[i].employee_name,
                    rule="休息不足",
                    details=f"两班间隔仅 {rest:.1f} 小时（要求 ≥{min_rest_hours}h）",
                    dates=[sorted_recs[i].date, sorted_recs[i + 1].date],
                    severity="critical",
                ))
    return violations


def check_consecutive_days(
    records: list[ShiftRecord], max_days: int = 6,
) -> list[Violation]:
    violations: list[Violation] = []
    by_emp = defaultdict(list)
    for r in records:
        by_emp[r.employee_id].append(r)

    for eid, recs in by_emp.items():
        dates = sorted(set(r.date for r in recs))
        if not dates:
            continue
        streak_start = dates[0]
        streak = [dates[0]]
        for i in range(1, len(dates)):
            if (dates[i] - dates[i - 1]).days == 1:
                streak.append(dates[i])
            else:
                if len(streak) > max_days:
                    violations.append(Violation(
                        employee_id=eid,
                        employee_name=next(r.employee_name for r in recs if r.date == streak[0]),
                        rule="连续上班超限",
                        details=f"连续上班 {len(streak)} 天（上限 {max_days} 天）",
                        dates=list(streak),
                        severity="warning",
                    ))
                streak_start = dates[i]
                streak = [dates[i]]
        if len(streak) > max_days:
            violations.append(Violation(
                employee_id=eid,
                employee_name=next(r.employee_name for r in recs if r.date == streak[0]),
                rule="连续上班超限",
                details=f"连续上班 {len(streak)} 天（上限 {max_days} 天）",
                dates=list(streak),
                severity="warning",
            ))
    return violations


def _iso_week(d: datetime.date) -> tuple[int, int]:
    iso = d.isocalendar()
    return (iso[0], iso[1])


def check_weekly_overtime(
    records: list[ShiftRecord], threshold: float = 40.0,
) -> list[Violation]:
    violations: list[Violation] = []
    by_emp_week = defaultdict(list)
    for r in records:
        by_emp_week[(r.employee_id, _iso_week(r.date))].append(r)

    for (eid, week), recs in by_emp_week.items():
        total = sum(r.hours for r in recs)
        if total > threshold:
            violations.append(Violation(
                employee_id=eid,
                employee_name=recs[0].employee_name,
                rule="周工时超标",
                details=f"第{week[1]}周总工时 {total:.1f}h（上限 {threshold}h）",
                dates=sorted(set(r.date for r in recs)),
                severity="warning",
            ))
    return violations


def check_monthly_overtime(
    records: list[ShiftRecord], threshold: float = 36.0,
    standard_weekly: float = 40.0,
) -> list[Violation]:
    violations: list[Violation] = []
    by_emp_month = defaultdict(list)
    for r in records:
        by_emp_month[(r.employee_id, r.date.year, r.date.month)].append(r)

    for (eid, y, m), recs in by_emp_month.items():
        total = sum(r.hours for r in recs)
        weeks_in_month = len(set(_iso_week(r.date) for r in recs))
        standard = weeks_in_month * standard_weekly
        overtime = total - standard
        if overtime > threshold:
            violations.append(Violation(
                employee_id=eid,
                employee_name=recs[0].employee_name,
                rule="月加班超标",
                details=f"{y}年{m}月加班 {overtime:.1f}h（上限 {threshold}h）",
                dates=sorted(set(r.date for r in recs)),
                severity="critical",
            ))
    return violations


def check_night_then_early(
    records: list[ShiftRecord],
    night_end_hour: int = 6,
    early_start_hour: int = 8,
) -> list[Violation]:
    violations: list[Violation] = []
    by_emp = defaultdict(list)
    for r in records:
        by_emp[r.employee_id].append(r)

    for eid, recs in by_emp.items():
        sorted_recs = sorted(recs, key=lambda x: _shift_start_dt(x))
        for i in range(len(sorted_recs) - 1):
            curr = sorted_recs[i]
            nxt = sorted_recs[i + 1]
            if curr.overnight:
                curr_end_dt = _shift_end_dt(curr)
                nxt_start_dt = _shift_start_dt(nxt)
                rest_hours = (nxt_start_dt - curr_end_dt).total_seconds() / 3600
                nxt_start = nxt.start_time or datetime.time(9, 0)
                if nxt_start.hour < early_start_hour:
                    violations.append(Violation(
                        employee_id=eid,
                        employee_name=curr.employee_name,
                        rule="夜班后接早班",
                        details=f"夜班 {(curr.end_time or datetime.time(5,0)).strftime('%H:%M')} 结束后次日 {nxt_start.strftime('%H:%M')} 上班，休息仅 {rest_hours:.1f}h",
                        dates=[curr.date, nxt.date],
                        severity="critical",
                    ))
    return violations


def check_special_employee_restrictions(
    records: list[ShiftRecord],
    employees: list[dict],
    rules: dict | None = None,
) -> list[Violation]:
    if rules is None:
        rules = load_rules()

    violations: list[Violation] = []
    emp_lookup = build_employee_lookup(employees)
    by_emp = defaultdict(list)
    for r in records:
        by_emp[r.employee_id].append(r)

    minor_forbidden = set(rules.get("minor_forbidden_shifts", ["C", "晚"]))
    pregnant_forbidden = set(rules.get("pregnant_forbidden_shifts", ["C", "晚"]))
    minor_max_hours = rules.get("minor_max_daily_hours", 8)
    pregnant_max_hours = rules.get("pregnant_max_daily_hours", 8)

    for eid, recs in by_emp.items():
        emp = emp_lookup.get("by_id", {}).get(eid)
        if emp is None:
            continue

        is_minor = emp.get("is_minor", False)
        is_pregnant = emp.get("is_pregnant", False)

        for r in recs:
            if is_minor:
                if r.shift_code in minor_forbidden:
                    violations.append(Violation(
                        employee_id=eid,
                        employee_name=r.employee_name,
                        rule="未成年禁止班次",
                        details=f"未成年员工排了 {r.shift_code} 班",
                        dates=[r.date],
                        severity="critical",
                    ))
                if r.hours > minor_max_hours:
                    violations.append(Violation(
                        employee_id=eid,
                        employee_name=r.employee_name,
                        rule="未成年超工时",
                        details=f"未成年日工时 {r.hours:.1f}h（上限 {minor_max_hours}h）",
                        dates=[r.date],
                        severity="critical",
                    ))

            if is_pregnant:
                if r.shift_code in pregnant_forbidden:
                    violations.append(Violation(
                        employee_id=eid,
                        employee_name=r.employee_name,
                        rule="孕期禁止班次",
                        details=f"孕期员工排了 {r.shift_code} 班",
                        dates=[r.date],
                        severity="critical",
                    ))
                if r.hours > pregnant_max_hours:
                    violations.append(Violation(
                        employee_id=eid,
                        employee_name=r.employee_name,
                        rule="孕期超工时",
                        details=f"孕期日工时 {r.hours:.1f}h（上限 {pregnant_max_hours}h）",
                        dates=[r.date],
                        severity="critical",
                    ))

    return violations


def run_all_checks(
    records: list[ShiftRecord],
    employees: list[dict] | None = None,
    rules: dict | None = None,
) -> list[Violation]:
    if employees is None:
        employees = load_employees()
    if rules is None:
        rules = load_rules()

    all_violations: list[Violation] = []
    all_violations.extend(check_cross_store_overlap(records))
    all_violations.extend(check_rest_insufficiency(records, rules.get("min_rest_hours", 11)))
    all_violations.extend(check_consecutive_days(records, rules.get("max_consecutive_days", 6)))
    all_violations.extend(check_weekly_overtime(records, rules.get("weekly_overtime_threshold", 40)))
    all_violations.extend(check_monthly_overtime(records, rules.get("monthly_overtime_threshold", 36)))
    all_violations.extend(check_night_then_early(
        records,
        rules.get("night_shift_end_hour", 6),
        rules.get("early_shift_start_hour", 8),
    ))
    all_violations.extend(check_special_employee_restrictions(records, employees, rules))

    return all_violations
