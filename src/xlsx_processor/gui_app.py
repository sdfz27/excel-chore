"""Desktop GUI for xlsx-processor: select file, workbooks, max column, output; run same export as CLI."""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from xlsx_processor.cli import export_employee_hours
from xlsx_processor.reader import read_workbook_names


def main() -> None:
    app = tk.Tk()
    app.title("xlsx-processor")
    app.minsize(420, 380)
    app.columnconfigure(0, weight=1)

    # State
    input_path = tk.StringVar(app, value="")
    sheet_names_list: list[str] = []
    sheet_vars: list[tk.BooleanVar] = []
    sheet_frame: tk.Frame | None = None
    max_col_var = tk.StringVar(app, value="AS")
    output_path_var = tk.StringVar(app, value="")
    status_var = tk.StringVar(app, value="")

    def on_browse_input() -> None:
        path = filedialog.askopenfilename(
            title="Select XLSX file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return
        input_path.set(path)
        status_var.set("")
        nonlocal sheet_vars, sheet_frame, sheet_names_list
        sheet_names_list = []
        for v in sheet_vars:
            v.set(False)
        sheet_vars.clear()
        if sheet_frame:
            sheet_frame.destroy()
        try:
            names = read_workbook_names(path)
            sheet_names_list = names
        except Exception as e:
            status_var.set(f"Error reading file: {e}")
            return
        sheet_frame = ttk.LabelFrame(inner, text="2. Workbooks to process", padding=6)
        sheet_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        sheet_frame.columnconfigure(0, weight=1)
        for i, name in enumerate(names):
            var = tk.BooleanVar(app, value=True)
            sheet_vars.append(var)
            cb = ttk.Checkbutton(sheet_frame, text=name, variable=var)
            cb.grid(row=i, column=0, sticky="w")
        inner.columnconfigure(0, weight=1)

    def on_browse_output() -> None:
        path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            output_path_var.set(path)

    def on_process() -> None:
        path_str = input_path.get().strip()
        if not path_str:
            messagebox.showwarning("Input required", "Please select an XLSX file.")
            return
        path = Path(path_str)
        if not path.exists():
            messagebox.showerror("Error", f"File not found: {path}")
            return
        selected = [s for s, v in zip(sheet_names_list, sheet_vars) if v.get()]
        if not selected and sheet_vars:
            messagebox.showwarning("Selection required", "Please select at least one workbook.")
            return
        if not selected:
            selected = None  # auto-detect
        max_col = max_col_var.get().strip() or None
        out_str = output_path_var.get().strip()
        if not out_str:
            messagebox.showwarning("Output required", "Please specify output location.")
            return
        out_path = Path(out_str)
        status_var.set("Processing...")
        app.update()
        try:
            n, err = export_employee_hours([path], selected, max_col, out_path)
            if err:
                status_var.set("")
                messagebox.showerror("Export failed", err)
            else:
                status_var.set(f"Done. Wrote {n} rows to {out_path.name}")
                messagebox.showinfo("Done", f"Exported {n} rows to\n{out_path}")
        except Exception as e:
            status_var.set("")
            messagebox.showerror("Error", str(e))

    # Layout
    inner = ttk.Frame(app, padding=12)
    inner.grid(row=0, column=0, sticky="nsew")
    inner.columnconfigure(0, weight=1)
    app.rowconfigure(0, weight=1)

    row = 0

    # 1. Select xlsx
    lf_input = ttk.LabelFrame(inner, text="1. Select XLSX to open", padding=6)
    lf_input.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 8))
    lf_input.columnconfigure(0, weight=1)
    ttk.Button(lf_input, text="Browse...", command=on_browse_input).grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Label(lf_input, textvariable=input_path, foreground="gray").grid(row=0, column=1, sticky="w")
    row += 1

    # 2. Workbooks checklist (populated after file selected)
    sheet_frame = ttk.LabelFrame(inner, text="2. Workbooks to process", padding=6)
    sheet_frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 8))
    sheet_frame.columnconfigure(0, weight=1)
    ttk.Label(sheet_frame, text="Select a file first.", foreground="gray").grid(row=0, column=0, sticky="w")
    row += 1

    # 3. Max column
    lf_col = ttk.LabelFrame(inner, text="3. Max column (e.g. AS)", padding=6)
    lf_col.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 8))
    lf_col.columnconfigure(0, weight=1)
    ttk.Entry(lf_col, textvariable=max_col_var, width=10).grid(row=0, column=0, sticky="w")
    row += 1

    # 4. Output location
    lf_out = ttk.LabelFrame(inner, text="4. Output location", padding=6)
    lf_out.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(0, 8))
    lf_out.columnconfigure(0, weight=1)
    ttk.Entry(lf_out, textvariable=output_path_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
    ttk.Button(lf_out, text="Browse...", command=on_browse_output).grid(row=0, column=1, sticky="w")
    row += 1

    # Process
    ttk.Button(inner, text="Process", command=on_process).grid(row=row, column=0, columnspan=2, pady=(8, 4))
    row += 1
    ttk.Label(inner, textvariable=status_var, foreground="green").grid(row=row, column=0, columnspan=2, sticky="w")

    app.mainloop()


if __name__ == "__main__":
    main()
