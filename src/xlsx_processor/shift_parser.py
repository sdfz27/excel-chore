"""Parse Shift workbook from Book1.xlsx into EmployeePage list using user.Employee/HourRecord/EmployeePage."""

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from xlsx_processor.user import Employee, EmployeePage, HourRecord

PAGE_MARKER = "SUPERVISOR SIGNATURE"
EMP_NUM_HEADER = "Emp #"
NAME_HEADER = "Name"


def _cell_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _special_code_str(v: Any) -> str:
    """Return special_code as string; treat None and 0 as empty (no "0" for missing values)."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    try:
        if float(s) == 0:
            return ""
    except ValueError:
        pass
    return s


def _find_header_row(rows: list[tuple[Any, ...]]) -> int:
    """Find 0-based row index that contains both 'Emp #' and 'Name'."""
    for i, row in enumerate(rows):
        if row is None:
            continue
        row_str = [_cell_str(v) for v in row]
        if EMP_NUM_HEADER in row_str and NAME_HEADER in row_str:
            return i
    raise ValueError("Shift sheet has no row with 'Emp #' and 'Name'")


def _record_name_from_block(header_row: tuple[Any, ...], start: int) -> str:
    """First non-empty value in header_row[start:start+3] as record name."""
    for j in range(3):
        if start + j < len(header_row):
            v = header_row[start + j]
            if v is not None and str(v).strip():
                return str(v).strip()
    return ""


def _record_code_from_block(code_row: tuple[Any, ...], start: int) -> str:
    """First non-empty value in code_row[start:start+3] as record code; if value contains newline, use the part after last newline (e.g. 'Downtime\\n9000' -> '9000')."""
    for j in range(3):
        if start + j < len(code_row):
            v = code_row[start + j]
            if v is not None:
                s = str(v).strip()
                if s:
                    if "\n" in s:
                        s = s.split("\n")[-1].strip()
                    return s
    return ""


def _max_column_to_index(value: str | int | None) -> int | None:
    """Convert -c/--max-column value (e.g. 'AS') to 1-based column index."""
    if value is None:
        return None
    if isinstance(value, int):
        return value if value >= 1 else None
    s = str(value).strip().upper()
    if not s:
        return None
    try:
        return column_index_from_string(s)
    except Exception:
        return None


def _is_page_break_row(row: tuple[Any, ...] | None) -> bool:
    if row is None:
        return False
    return any(
        v is not None and PAGE_MARKER in str(v)
        for v in row
    )


def parse_shift_to_employee_pages(
    path: str | Path,
    sheet_name: str | None = None,
    max_column: str | int | None = None,
) -> list[EmployeePage]:
    """
    Parse the Shift workbook in the given xlsx into a list of EmployeePage.

    - employee_id from "Emp #" column, name from "Name" column.
    - Rows containing "SUPERVISOR SIGNATURE" start a new page.
    - After "Name", every 3 columns form one HourRecord (rt, t15, r20);
      record_name is taken from the header row above the Emp # row (row index - 2).
    - If max_column is set (e.g. "AS"), only columns up to that column are read.

    Args:
        path: Path to the xlsx file.
        sheet_name: Workbook name containing the shift data; if None, use first sheet
            whose name contains "Shift".
        max_column: Optional max column letter or 1-based index; only read data up to this column.

    Returns:
        List of EmployeePage, one per "page" (separated by SUPERVISOR SIGNATURE rows).
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name is not None:
            ws = wb[sheet_name]
        else:
            name = next((s for s in wb.sheetnames if "Shift" in s), None)
            if name is None:
                raise ValueError(f"未找到名称包含 'Shift' 的工作表，当前: {wb.sheetnames}")
            ws = wb[name]

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
    finally:
        wb.close()

    header_row_idx = _find_header_row(rows)
    emp_col = rows[header_row_idx].index(EMP_NUM_HEADER)
    name_col = rows[header_row_idx].index(NAME_HEADER)

    # Record names: row two above header (e.g. "Selecting", "Loading"); record_code: row above that (e.g. "Kronos Hrs", "2000", "9000")
    record_name_row_idx = max(0, header_row_idx - 2)
    record_name_row = rows[record_name_row_idx]
    record_code_row_idx = max(0, record_name_row_idx - 1)
    record_code_row = rows[record_code_row_idx]

    max_col_1based = _max_column_to_index(max_column)

    pages: list[EmployeePage] = []
    current_employees: list[Employee] = []

    for i in range(header_row_idx + 1, len(rows)):
        row = rows[i]
        if row is None:
            continue
        if _is_page_break_row(row):
            if current_employees:
                pages.append(EmployeePage(employees=list(current_employees)))
                current_employees = []
            continue

        emp_id_val = row[emp_col] if emp_col < len(row) else None
        name_val = row[name_col] if name_col < len(row) else None
        if emp_id_val is None and name_val is None:
            continue
        employee_id = _cell_str(emp_id_val)
        name = _cell_str(name_val)
        special_col = emp_col + 1
        special_val = row[special_col] if special_col < len(row) else None
        special_code = _special_code_str(special_val)

        hour_records: list[HourRecord] = []
        start = name_col + 1
        while start + 2 < len(row):
            # 0-based start -> 1-based column of block end = start + 3; stop if beyond max_column
            if max_col_1based is not None and (start + 3) > max_col_1based:
                break
            record_code = _record_code_from_block(record_code_row, start)
            record_name = _record_name_from_block(record_name_row, start)
            rt = _cell_str(row[start] if start < len(row) else None)
            t15 = _cell_str(row[start + 1] if start + 1 < len(row) else None)
            r20 = _cell_str(row[start + 2] if start + 2 < len(row) else None)
            if not record_name:
                record_name = f"Col{start}"
            hour_records.append(
                HourRecord(
                    record_code=record_code,
                    record_name=record_name,
                    rt=rt,
                    t15=t15,
                    r20=r20,
                )
            )
            start += 3

        current_employees.append(
            Employee(
                employee_id=employee_id,
                name=name,
                special_code=special_code,
                hour_records=hour_records,
            )
        )

    if current_employees:
        pages.append(EmployeePage(employees=list(current_employees)))

    return pages
