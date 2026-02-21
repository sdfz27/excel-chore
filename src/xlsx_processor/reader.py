"""读取 xlsx 文件及其中各 workbook（工作表）名称与行数据。"""

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


def read_workbook_names(path: str | Path) -> list[str]:
    """
    读取指定 xlsx 文件，返回其中所有工作表（sheet）的名称列表。

    Args:
        path: xlsx 文件路径。

    Returns:
        工作表名称列表，按在文件中的顺序。

    Raises:
        FileNotFoundError: 文件不存在。
        ValueError: 文件不是有效的 xlsx。
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(f"不支持的文件格式: {path.suffix}，请使用 .xlsx 或 .xlsm")

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def read_workbook_names_from_paths(paths: list[str | Path]) -> dict[str, list[str]]:
    """
    批量读取多个 xlsx 文件，返回「文件路径 -> 工作表名称列表」的映射。

    Args:
        paths: xlsx 文件路径列表。

    Returns:
        { 文件路径: [工作表名1, 工作表名2, ...], ... }
    """
    result: dict[str, list[str]] = {}
    for p in paths:
        key = str(Path(p).resolve())
        result[key] = read_workbook_names(p)
    return result


def _parse_max_column(value: str | int | None) -> int | None:
    """将最大列参数转为 1-based 列索引；支持字母（如 AS）或数字。"""
    if value is None:
        return None
    if isinstance(value, int):
        if value < 1:
            raise ValueError(f"列索引必须 ≥ 1，得到: {value}")
        return value
    s = str(value).strip().upper()
    if not s:
        return None
    try:
        return column_index_from_string(s)
    except Exception as e:
        raise ValueError(f"无效的最大列参数: {value!r}，应为字母（如 A、AS）或正整数") from e


def read_sheet_rows(
    path: str | Path,
    sheet_names: list[str] | None = None,
    max_column: str | int | None = None,
) -> dict[str, list[list[Any]]]:
    """
    读取指定 xlsx 文件中若干工作表的行数据。

    Args:
        path: xlsx 文件路径。
        sheet_names: 要读取的工作表名称列表；为 None 时读取所有工作表。
        max_column: 最大列，仅读取从第 1 列到该列的数据。可为 Excel 列字母（如 "AS"）
            或 1-based 列号；为 None 时读取整行。

    Returns:
        { 工作表名: [ [cell1, cell2, ...], ... ], ... }，每行为一列表的单元格值。

    Raises:
        FileNotFoundError: 文件不存在。
        ValueError: 文件格式不支持、指定工作表不存在或 max_column 无效。
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {path}")
    if path.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError(f"不支持的文件格式: {path.suffix}，请使用 .xlsx 或 .xlsm")

    max_col_index = _parse_max_column(max_column)

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        available = set(wb.sheetnames)
        if sheet_names is not None:
            missing = set(sheet_names) - available
            if missing:
                raise ValueError(f"工作表不存在: {missing}，文件中仅有: {sorted(available)}")
            to_read = list(sheet_names)
        else:
            to_read = wb.sheetnames

        result: dict[str, list[list[Any]]] = {}
        for name in to_read:
            ws = wb[name]
            kwargs = {} if max_col_index is None else {"max_col": max_col_index}
            rows = [[cell.value for cell in row] for row in ws.iter_rows(**kwargs)]
            result[name] = rows
        return result
    finally:
        wb.close()
