from __future__ import annotations

import datetime

import pytest

from shift_checker.cost import (
    ShiftCost,
    compute_shift_costs,
    summarize_by_employee,
    summarize_by_store,
)
from shift_checker.parser import ShiftRecord


def _rec(
    eid="E001",
    name="张三",
    store="旗舰店",
    date=None,
    shift="A",
    start=None,
    end=None,
    hours=8.0,
    overnight=False,
) -> ShiftRecord:
    return ShiftRecord(
        employee_id=eid,
        employee_name=name,
        store=store,
        date=date or datetime.date(2025, 1, 6),
        shift_code=shift,
        start_time=start or datetime.time(9, 0),
        end_time=end or datetime.time(17, 0),
        hours=hours,
        overnight=overnight,
        source_file="test.xlsx",
        row_number=2,
    )


class TestComputeShiftCosts:
    def test_weekday_base_cost(self, employees, holidays, cost_rates):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, 6), hours=8.0)]
        costs = compute_shift_costs(recs, employees, holidays, cost_rates)
        assert len(costs) == 1
        c = costs[0]
        assert c.day_type == "weekday"
        assert c.multiplier == 1.0
        assert c.hourly_rate == 25.0
        assert c.base_cost == 8.0 * 25.0 * 1.0
        assert c.total_cost == c.base_cost

    def test_weekend_multiplier(self, employees, holidays, cost_rates):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, 11), hours=8.0)]
        costs = compute_shift_costs(recs, employees, holidays, cost_rates)
        assert len(costs) == 1
        c = costs[0]
        assert c.day_type == "weekend"
        assert c.multiplier == 1.5
        assert c.base_cost == 8.0 * 25.0 * 1.5

    def test_holiday_multiplier(self, employees, holidays, cost_rates):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, 1), hours=8.0)]
        costs = compute_shift_costs(recs, employees, holidays, cost_rates)
        assert len(costs) == 1
        c = costs[0]
        assert c.day_type == "holiday"
        assert c.multiplier == 3.0
        assert c.base_cost == 8.0 * 25.0 * 3.0

    def test_overtime_cost_separated(self, employees, holidays, cost_rates):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, 6), hours=10.0)]
        costs = compute_shift_costs(recs, employees, holidays, cost_rates)
        c = costs[0]
        assert c.overtime_hours > 0
        assert c.overtime_cost > 0
        assert c.total_cost == c.base_cost + c.overtime_cost


class TestSummarizeByEmployee:
    def test_summary(self, employees, holidays, cost_rates):
        recs = [
            _rec(eid="E001", date=datetime.date(2025, 1, 6), hours=8.0),
            _rec(eid="E001", date=datetime.date(2025, 1, 7), hours=8.0),
        ]
        costs = compute_shift_costs(recs, employees, holidays, cost_rates)
        summaries = summarize_by_employee(costs)
        assert len(summaries) == 1
        s = summaries[0]
        assert s.employee_id == "E001"
        assert s.total_hours == 16.0
        assert s.total_cost > 0


class TestSummarizeByStore:
    def test_summary(self, employees, holidays, cost_rates):
        recs = [
            _rec(eid="E001", store="旗舰店", hours=8.0),
            _rec(eid="E002", store="旗舰店", hours=8.0),
        ]
        costs = compute_shift_costs(recs, employees, holidays, cost_rates)
        summaries = summarize_by_store(costs)
        assert len(summaries) == 1
        assert summaries[0].store == "旗舰店"
        assert summaries[0].employee_count == 2
        assert summaries[0].total_hours == 16.0
