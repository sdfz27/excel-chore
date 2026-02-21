"""命令行入口：默认解析 Shift 为 EmployeePage 并打印；可导出非空 HourRecord 到 xlsx，或按工作表输出原始行。"""

import argparse
import sys
from pathlib import Path

from openpyxl import Workbook

from xlsx_processor.reader import (
    read_workbook_names_from_paths,
    read_sheet_rows,
)
from xlsx_processor.shift_parser import parse_shift_to_employee_pages
from xlsx_processor.user import HourRecord


def main() -> None:
    parser = argparse.ArgumentParser(
        description="读取 xlsx，默认解析 Shift 并打印 EmployeePage；可用 -o 导出非空 HourRecord 到 xlsx。"
    )
    parser.add_argument(
        "paths",
        nargs="+",
        type=Path,
        help="一个或多个 xlsx 文件路径",
    )
    parser.add_argument(
        "-o",
        "--output",
        metavar="XLSX",
        type=Path,
        help="将非全空的 HourRecord 写入此 xlsx（每行：员工信息 + record_name, rt, t15, r20）",
    )
    parser.add_argument(
        "-s",
        "--sheets",
        nargs="*",
        metavar="SHEET",
        help="指定后输出这些工作表的每一行，而非解析 EmployeePage",
    )
    parser.add_argument(
        "-c",
        "--max-column",
        metavar="COL",
        help="最大列（Excel 列字母）；仅与 -s 同用",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="减少输出（打印时只输出员工 id/name；与 -o 同用时仍会写入文件）",
    )
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="仅列出各文件中的工作表名，不解析 EmployeePage",
    )
    args = parser.parse_args()

    try:
        if args.list_sheets:
            _print_workbook_names(args.paths, args.quiet)
        elif args.sheets is not None and len(args.sheets) > 0:
            _print_sheet_rows(
                args.paths, args.sheets, args.quiet, args.max_column
            )
        else:
            _run_employee_default(args.paths, args.quiet, args.output)
    except (FileNotFoundError, ValueError) as e:
        print(e, file=sys.stderr)
        sys.exit(1)
    return None


def _is_hour_record_empty(hr: HourRecord) -> bool:
    return not (hr.rt or hr.t15 or hr.r20)


def _run_employee_default(
    paths: list[Path],
    quiet: bool,
    output_path: Path | None,
) -> None:
    """Parse Shift in each path; print EmployeePage and optionally write non-empty HourRecords to output xlsx."""
    all_rows: list[tuple[str, str, str, str, str, str, str]] = []
    for path in paths:
        try:
            pages = parse_shift_to_employee_pages(path)
        except (FileNotFoundError, ValueError) as e:
            print(f"{path}: {e}", file=sys.stderr)
            continue
        if not output_path:
            if not quiet:
                print(f"文件: {path.resolve()}")
            for page_idx, page in enumerate(pages):
                if not quiet:
                    print(f"  --- EmployeePage {page_idx + 1} ---")
                for emp in page.employees:
                    if quiet:
                        print(f"{emp.employee_id}\t{emp.name}")
                    else:
                        print(f"    {emp.employee_id}  {emp.name}")
                        for hr in emp.hour_records:
                            print(f"      [{hr.record_code}] {hr.record_name}: rt={hr.rt} t15={hr.t15} r20={hr.r20}")
                if not quiet:
                    print()
            if not quiet and pages:
                print()
        if output_path:
            for page in pages:
                for emp in page.employees:
                    for hr in emp.hour_records:
                        if not _is_hour_record_empty(hr):
                            all_rows.append(
                                (emp.employee_id, emp.name, hr.record_code, hr.record_name, hr.rt, hr.t15, hr.r20)
                            )
    if output_path:
        _write_employee_hours_xlsx(output_path, all_rows)
        if not quiet:
            n = len(all_rows)
            print(f"已写入 {n} 行到 {output_path.resolve()}", file=sys.stderr)


def _write_employee_hours_xlsx(
    path: Path,
    rows: list[tuple[str, str, str, str, str, str, str]],
) -> None:
    """Write rows (employee_id, name, record_code, record_name, rt, t15, r20) to a new xlsx."""
    wb = Workbook()
    ws = wb.active
    ws.title = "EmployeeHours"
    ws.append(["Employee ID", "Name", "Record Code", "Record Name", "RT", "T15", "R20"])
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def _print_workbook_names(paths: list[Path], quiet: bool) -> None:
    mapping = read_workbook_names_from_paths(paths)
    for file_path, names in mapping.items():
        if not quiet:
            print(f"文件: {file_path}")
            print(f"  工作表: {', '.join(names)}")
        else:
            for name in names:
                print(name)


def _print_sheet_rows(
    paths: list[Path],
    sheet_names: list[str],
    quiet: bool,
    max_column: str | None = None,
) -> None:
    for path in paths:
        data = read_sheet_rows(path, sheet_names, max_column=max_column)
        for sheet_name, rows in data.items():
            if not quiet:
                print(f"文件: {path.resolve()}")
                print(f"  工作表: {sheet_name}")
            for row in rows:
                print("\t".join(str(c) if c is not None else "" for c in row))
            if not quiet:
                print()


if __name__ == "__main__":
    main()
