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

### 桌面 GUI

```bash
uv run xlsx-processor-gui
```

- 选择要打开的 xlsx 文件
- 在清单中勾选要处理的工作表（workbook）
- 可选：指定最大列（如 AS）
- 指定输出文件路径
- 点击 Process 执行导出（逻辑与 CLI `-o` 一致）

### 打包为 Windows 桌面 exe

使用 [PyInstaller](https://pyinstaller.org/) 将 GUI 打成单个 Windows 可执行文件（无需安装 Python 即可运行）。

**1. 安装打包依赖（可选）**

```bash
uv sync --extra build-exe
# 或
uv pip install pyinstaller
```

**2. 在项目根目录执行打包**

在 `c:\projects\python\python-suji` 下执行（保证 `src` 和 `scripts` 存在）：

```bash
uv run pyinstaller --onefile --windowed -n xlsx-processor-gui --paths=src scripts/run_gui.py
```

- `--onefile`：生成单个 .exe（否则是一目录 + exe）
- `--windowed`：无控制台窗口（GUI 应用）
- `-n xlsx-processor-gui`：输出 exe 名称
- `--paths=src`：让 PyInstaller 找到 `xlsx_processor` 包
- `scripts/run_gui.py`：GUI 入口脚本

**3. 获取 exe**

打包完成后，可执行文件在：

```
dist/xlsx-processor-gui.exe
```

将 `dist\xlsx-processor-gui.exe` 复制到任意位置即可在 Windows 上运行，无需本机安装 Python。

**可选：指定图标**

若有 `.ico` 文件，可加上：

```bash
uv run pyinstaller --onefile --windowed -n xlsx-processor-gui --paths=src --icon=app.ico scripts/run_gui.py
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
