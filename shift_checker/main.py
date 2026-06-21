from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from shift_checker.compliance import run_all_checks
from shift_checker.cost import compute_shift_costs, summarize_by_employee, summarize_by_store
from shift_checker.output import (
    generate_markdown_report,
    write_cost_summary,
    write_merged_table,
    write_violations,
)
from shift_checker.parser import parse_all_workbooks


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="排班合规检查与成本核算工具",
    )
    parser.add_argument(
        "input_dir",
        type=str,
        help="包含门店排班 Excel 文件的目录",
    )
    parser.add_argument(
        "-o", "--output-dir",
        type=str,
        default=".",
        help="输出目录（默认当前目录）",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="显示详细日志",
    )

    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    input_dir = Path(args.input_dir)
    if not input_dir.is_dir():
        print(f"错误: 输入目录不存在: {input_dir}", file=sys.stderr)
        return 1

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"正在解析 {input_dir} 中的排班文件...")
    records = parse_all_workbooks(input_dir)
    print(f"共解析到 {len(records)} 条排班记录")

    if not records:
        print("未找到有效排班记录，退出")
        return 0

    print("\n正在执行合规检查...")
    violations = run_all_checks(records)
    print(f"发现 {len(violations)} 条违规")
    for v in violations:
        icon = "🔴" if v.severity == "critical" else "🟡"
        print(f"  {icon} {v.employee_name}({v.employee_id}): {v.rule} - {v.details}")

    print("\n正在核算成本...")
    shift_costs = compute_shift_costs(records)
    emp_summaries = summarize_by_employee(shift_costs)
    store_summaries = summarize_by_store(shift_costs)
    total_cost = sum(s.total_cost for s in emp_summaries)
    print(f"总人力成本: ¥{total_cost:,.2f}")

    merged_path = output_dir / "合并总表.xlsx"
    violation_path = output_dir / "违规清单.xlsx"
    cost_path = output_dir / "成本汇总.xlsx"
    md_path = output_dir / "排班分析报告.md"

    write_merged_table(records, merged_path)
    write_violations(violations, violation_path)
    write_cost_summary(emp_summaries, store_summaries, cost_path)
    generate_markdown_report(records, violations, emp_summaries, store_summaries, md_path)

    print(f"\n输出文件:")
    print(f"  合并总表: {merged_path}")
    print(f"  违规清单: {violation_path}")
    print(f"  成本汇总: {cost_path}")
    print(f"  分析报告: {md_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
