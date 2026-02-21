# xlsx-processor

使用 Python 读取 xlsx 文件：可列出 workbook（工作表）名称，或指定工作表并列出其中每一行；支持多个文件。

## 环境要求

- Python ≥ 3.10
- [uv](https://docs.astral.sh/uv/)（推荐用于依赖与虚拟环境管理）

## 项目结构（src layout）

```
python-suji/
├── pyproject.toml
├── README.md
└── src/
    └── xlsx_processor/
        ├── __init__.py
        ├── reader.py    # 读取 xlsx、获取工作表名
        └── cli.py       # 命令行入口
```

## 安装与运行

```bash
# 进入项目目录
cd c:\projects\python\python-suji

# 使用 uv 创建虚拟环境并安装依赖
uv sync

# 默认：解析 Shift 工作表并打印所有 EmployeePage
uv run xlsx-processor path/to/file.xlsx
uv run xlsx-processor file1.xlsx file2.xlsx
uv run xlsx-processor path/to/file.xlsx -q   # 仅打印 employee_id 与 name

# 将非空 HourRecord 导出到 xlsx（每行：Employee ID, Name, Record Name, RT, T15, R20）
uv run xlsx-processor path/to/file.xlsx -o output.xlsx
uv run xlsx-processor file1.xlsx file2.xlsx -o combined.xlsx

# 仅列出各文件的工作表名
uv run xlsx-processor path/to/file.xlsx --list-sheets

# 指定工作表并输出每一行（先写文件路径，再 -s 表名）
uv run xlsx-processor path/to/file.xlsx -s Sheet1 销售表 库存表
uv run xlsx-processor file1.xlsx file2.xlsx -s Sheet1 Sheet2

# 仅输出行数据（不输出文件名/表名）
uv run xlsx-processor -q path/to/file.xlsx -s Sheet1

# 限制最大列：每行只包含第 1 列到该列（如 AS）的数据
uv run xlsx-processor path/to/file.xlsx -s Sheet1 -c AS
uv run xlsx-processor path/to/file.xlsx -s Sheet1 --max-column AS
```

## 在代码中使用

```python
from xlsx_processor import read_workbook_names, read_sheet_rows
from xlsx_processor.reader import read_workbook_names_from_paths

# 单个文件：获取工作表名列表
names = read_workbook_names("data.xlsx")
print(names)  # ['Sheet1', 'Sheet2', ...]

# 多个文件：获取「文件路径 -> 工作表名列表」字典
mapping = read_workbook_names_from_paths(["a.xlsx", "b.xlsx"])
for path, sheet_names in mapping.items():
    print(f"{path}: {sheet_names}")

# 读取指定工作表的每一行：返回 { 表名: [ [cell, ...], ... ] }
rows_by_sheet = read_sheet_rows("data.xlsx", sheet_names=["Sheet1", "销售表"])
for sheet_name, rows in rows_by_sheet.items():
    for row in rows:
        print(row)  # 每行为一列表的单元格值

# 只读到某列为止（如到 AS 列）：max_column 可为列字母或 1-based 列号
rows_by_sheet = read_sheet_rows("data.xlsx", sheet_names=["Sheet1"], max_column="AS")
```

## 依赖

- [openpyxl](https://openpyxl.readthedocs.io/)：读写 .xlsx / .xlsm
