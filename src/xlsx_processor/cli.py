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
        help="与 -o 同用时指定要解析的 Shift 工作表名（如 Shift1）；否则输出这些工作表的原始行",
    )
    parser.add_argument(
        "-c",
        "--max-column",
        metavar="COL",
        help="最大列（Excel 列字母，如 AS）；解析/输出时只包含到该列的数据（与 -s 或默认 employee 均可用）",
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
        elif args.output is not None:
            # -o: always parse Shift and write; -s can specify one or more sheets to parse (e.g. Shift1 Shift2); -c limits columns
            shift_sheets = args.sheets if args.sheets else None
            _run_employee_default(args.paths, args.quiet, args.output, shift_sheets, args.max_column)
        elif args.sheets is not None and len(args.sheets) > 0:
            _print_sheet_rows(
                args.paths, args.sheets, args.quiet, args.max_column
            )
        else:
            _run_employee_default(args.paths, args.quiet, None, None, args.max_column)
    except (FileNotFoundError, ValueError) as e:
        print(e, file=sys.stderr)
        sys.exit(1)
    return None


def _is_value_empty(s: str) -> bool:
    """Treat None/empty string and '0' / '0.0' as empty."""
    if not s or not s.strip():
        return True
    try:
        return float(s.strip()) == 0
    except ValueError:
        return False


def _is_valid_hour_value(s: str) -> bool:
    """True if s is empty/whitespace or a valid number (e.g. 8, 2.5, 0). Rejects text like 'Selecting', 'Reg HRS'."""
    if not s or not s.strip():
        return True
    try:
        float(s.strip())
        return True
    except ValueError:
        return False


def _is_hour_record_empty(hr: HourRecord) -> bool:
    """Filter out if rt, t15, and r20 are all empty or zero."""
    return _is_value_empty(hr.rt) and _is_value_empty(hr.t15) and _is_value_empty(hr.r20)


def _is_hour_record_valid(hr: HourRecord) -> bool:
    """Filter out if any of rt, t15, r20 contains non-numeric text (e.g. header/activity names)."""
    return _is_valid_hour_value(hr.rt) and _is_valid_hour_value(hr.t15) and _is_valid_hour_value(hr.r20)


def _run_employee_default(
    paths: list[Path],
    quiet: bool,
    output_path: Path | None,
    shift_sheet_names: list[str] | None = None,
    max_column: str | None = None,
) -> None:
    """Parse Shift in each path; print EmployeePage and optionally write non-empty HourRecords to output xlsx.
    When shift_sheet_names is set (e.g. from -s Shift1 Shift2), parse each of those sheets per path; when None, auto-detect one Shift sheet.
    When max_column is set (e.g. AS), only read data up to that column.
    """
    if output_path:
        n, err = export_employee_hours(paths, shift_sheet_names, max_column, output_path)
        if err:
            print(err, file=sys.stderr)
        elif not quiet:
            print(f"已写入 {n} 行到 {output_path.resolve()}", file=sys.stderr)
        return

    for path in paths:
        sheets_to_parse: list[str | None] = (
            list(shift_sheet_names) if shift_sheet_names else [None]
        )
        for sheet_name in sheets_to_parse:
            try:
                pages = parse_shift_to_employee_pages(
                    path, sheet_name=sheet_name, max_column=max_column
                )
            except (FileNotFoundError, ValueError) as e:
                print(f"{path}" + (f" [{sheet_name}]" if sheet_name else "") + f": {e}", file=sys.stderr)
                continue
            if not quiet:
                print(f"文件: {path.resolve()}" + (f"  工作表: {sheet_name}" if sheet_name else ""))
            for page_idx, page in enumerate(pages):
                if not quiet:
                    print(f"  --- EmployeePage {page_idx + 1} ---")
                for emp in page.employees:
                    if not (emp.employee_id or "").strip():
                        continue
                    if quiet:
                        print(f"{emp.employee_id}\t{emp.name}\t{emp.special_code}")
                    else:
                        print(f"    {emp.employee_id}  {emp.name}  [{emp.special_code}]")
                        for hr in emp.hour_records:
                            if _is_hour_record_empty(hr) or not _is_hour_record_valid(hr):
                                continue
                            print(f"      [{hr.record_code}] {hr.record_name}: rt={hr.rt} t15={hr.t15} r20={hr.r20}")
                if not quiet:
                    print()
            if not quiet and pages:
                print()


def _write_employee_hours_xlsx(
    path: Path,
    rows: list[tuple[str, str, str, str, str, str, str, str]],
) -> None:
    """Write rows (employee_id, name, special_code, record_code, record_name, rt, t15, r20) to a new xlsx."""
    wb = Workbook()
    ws = wb.active
    ws.title = "EmployeeHours"
    ws.append(["Employee ID", "Name", "Special Code", "Record Code", "Record Name", "RT", "T15", "R20"])
    for r in rows:
        ws.append(list(r))
    wb.save(path)


def export_employee_hours(
    paths: list[Path],
    sheet_names: list[str] | None,
    max_column: str | None,
    output_path: Path,
) -> tuple[int, str | None]:
    """Parse selected sheets and write non-empty, valid HourRecords to output xlsx.
    Returns (rows_written, None) on success, or (0, error_message) on failure.
    Used by both CLI and GUI.
    """
    all_rows: list[tuple[str, str, str, str, str, str, str, str]] = []
    for path in paths:
        sheets_to_parse: list[str | None] = list(sheet_names) if sheet_names else [None]
        for sheet_name in sheets_to_parse:
            try:
                pages = parse_shift_to_employee_pages(
                    path, sheet_name=sheet_name, max_column=max_column
                )
            except (FileNotFoundError, ValueError) as e:
                return 0, str(e)
            for page in pages:
                for emp in page.employees:
                    if not (emp.employee_id or "").strip():
                        continue
                    for hr in emp.hour_records:
                        if not _is_hour_record_empty(hr) and _is_hour_record_valid(hr):
                            all_rows.append(
                                (
                                    emp.employee_id,
                                    emp.name,
                                    emp.special_code,
                                    hr.record_code,
                                    hr.record_name,
                                    hr.rt,
                                    hr.t15,
                                    hr.r20,
                                )
                            )
    try:
        _write_employee_hours_xlsx(output_path, all_rows)
        return len(all_rows), None
    except Exception as e:
        return 0, str(e)


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
