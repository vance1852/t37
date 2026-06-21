from __future__ import annotations

import datetime

import pytest

from shift_checker.compliance import (
    Violation,
    check_consecutive_days,
    check_cross_store_overlap,
    check_night_then_early,
    check_rest_insufficiency,
    check_special_employee_restrictions,
    check_weekly_overtime,
    run_all_checks,
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


class TestCrossStoreOverlap:
    def test_no_overlap(self):
        recs = [
            _rec(eid="E001", store="旗舰店", date=datetime.date(2025, 1, 6)),
            _rec(eid="E002", store="分店A", date=datetime.date(2025, 1, 6)),
        ]
        violations = check_cross_store_overlap(recs)
        assert len(violations) == 0

    def test_overlap_detected(self):
        recs = [
            _rec(eid="E001", store="旗舰店", date=datetime.date(2025, 1, 6)),
            _rec(eid="E001", store="分店A", date=datetime.date(2025, 1, 6)),
        ]
        violations = check_cross_store_overlap(recs)
        assert len(violations) == 1
        assert violations[0].rule == "跨店重叠排班"
        assert violations[0].severity == "critical"


class TestRestInsufficiency:
    def test_sufficient_rest(self):
        recs = [
            _rec(eid="E001", date=datetime.date(2025, 1, 6),
                 end=datetime.time(17, 0), overnight=False),
            _rec(eid="E001", date=datetime.date(2025, 1, 7),
                 start=datetime.time(9, 0), end=datetime.time(17, 0)),
        ]
        violations = check_rest_insufficiency(recs, min_rest_hours=11)
        assert len(violations) == 0

    def test_insufficient_rest(self):
        recs = [
            _rec(eid="E001", date=datetime.date(2025, 1, 6),
                 start=datetime.time(21, 0), end=datetime.time(5, 0),
                 shift="C", overnight=True, hours=8.0),
            _rec(eid="E001", date=datetime.date(2025, 1, 7),
                 start=datetime.time(8, 0), end=datetime.time(16, 0),
                 hours=8.0),
        ]
        violations = check_rest_insufficiency(recs, min_rest_hours=11)
        assert len(violations) == 1
        assert violations[0].rule == "休息不足"


class TestConsecutiveDays:
    def test_within_limit(self):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, i)) for i in range(1, 5)]
        violations = check_consecutive_days(recs, max_days=6)
        assert len(violations) == 0

    def test_exceeds_limit(self):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, i)) for i in range(1, 8)]
        violations = check_consecutive_days(recs, max_days=6)
        assert len(violations) == 1
        assert violations[0].rule == "连续上班超限"


class TestWeeklyOvertime:
    def test_within_threshold(self):
        recs = [_rec(eid="E001", date=datetime.date(2025, 1, 6), hours=8.0)]
        violations = check_weekly_overtime(recs, threshold=40)
        assert len(violations) == 0

    def test_exceeds_threshold(self):
        recs = [
            _rec(eid="E001", date=datetime.date(2025, 1, 6), hours=12.0),
            _rec(eid="E001", date=datetime.date(2025, 1, 7), hours=12.0),
            _rec(eid="E001", date=datetime.date(2025, 1, 8), hours=12.0),
            _rec(eid="E001", date=datetime.date(2025, 1, 9), hours=12.0),
        ]
        violations = check_weekly_overtime(recs, threshold=40)
        assert len(violations) == 1
        assert violations[0].rule == "周工时超标"


class TestNightThenEarly:
    def test_no_issue(self):
        recs = [
            _rec(eid="E001", date=datetime.date(2025, 1, 6),
                 shift="C", start=datetime.time(21, 0), end=datetime.time(5, 0),
                 overnight=True),
            _rec(eid="E001", date=datetime.date(2025, 1, 7),
                 start=datetime.time(14, 0), end=datetime.time(22, 0)),
        ]
        violations = check_night_then_early(recs)
        assert len(violations) == 0

    def test_night_then_early_detected(self):
        recs = [
            _rec(eid="E001", date=datetime.date(2025, 1, 6),
                 shift="C", start=datetime.time(21, 0), end=datetime.time(5, 0),
                 overnight=True),
            _rec(eid="E001", date=datetime.date(2025, 1, 7),
                 start=datetime.time(7, 0), end=datetime.time(15, 0)),
        ]
        violations = check_night_then_early(recs)
        assert len(violations) == 1
        assert violations[0].rule == "夜班后接早班"


class TestSpecialEmployeeRestrictions:
    def test_minor_night_shift_forbidden(self, employees, rules):
        recs = [
            _rec(eid="E004", name="赵六", shift="C",
                 start=datetime.time(21, 0), end=datetime.time(5, 0),
                 overnight=True, hours=8.0),
        ]
        violations = check_special_employee_restrictions(recs, employees, rules)
        assert any(v.rule == "未成年禁止班次" for v in violations)

    def test_pregnant_night_shift_forbidden(self, employees, rules):
        recs = [
            _rec(eid="E005", name="孙七", shift="晚",
                 start=datetime.time(21, 0), end=datetime.time(5, 0),
                 overnight=True, hours=8.0),
        ]
        violations = check_special_employee_restrictions(recs, employees, rules)
        assert any(v.rule == "孕期禁止班次" for v in violations)

    def test_normal_employee_no_restriction(self, employees, rules):
        recs = [
            _rec(eid="E001", name="张三", shift="C",
                 start=datetime.time(21, 0), end=datetime.time(5, 0),
                 overnight=True, hours=8.0),
        ]
        violations = check_special_employee_restrictions(recs, employees, rules)
        minor_v = [v for v in violations if "未成年" in v.rule]
        pregnant_v = [v for v in violations if "孕期" in v.rule]
        assert len(minor_v) == 0
        assert len(pregnant_v) == 0


class TestRunAllChecks:
    def test_returns_violations(self, employees, rules):
        recs = [
            _rec(eid="E001", store="旗舰店", date=datetime.date(2025, 1, 6)),
            _rec(eid="E001", store="分店A", date=datetime.date(2025, 1, 6)),
            _rec(eid="E004", name="赵六", shift="C",
                 start=datetime.time(21, 0), end=datetime.time(5, 0),
                 overnight=True, hours=8.0),
        ]
        violations = run_all_checks(recs, employees, rules)
        rules_found = {v.rule for v in violations}
        assert "跨店重叠排班" in rules_found
