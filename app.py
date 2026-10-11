"""Windows 11 Excel 自动整理工具的中文图形界面。"""

from pathlib import Path
from zipfile import BadZipFile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from organizer import merge_workbooks, read_header


class OrganizerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 自动整理工具")
        self.geometry("700x500")
        self.minsize(620, 440)
        self.paths: list[Path] = []
        self.sort_column = tk.StringVar()
        self.descending = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="请选择要合并的 .xlsx 文件。")
        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="1. 选择 Excel 文件（读取每个文件的第一个工作表）").pack(anchor="w")

        list_frame = ttk.Frame(frame)
        list_frame.pack(fill="both", expand=True, pady=(8, 6))
        self.file_list = tk.Listbox(list_frame, selectmode="extended")
        self.file_list.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        scrollbar.pack(side="right", fill="y")
        self.file_list.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="添加文件", command=self.add_files).pack(side="left")
        ttk.Button(buttons, text="移除选中", command=self.remove_selected).pack(side="left", padx=8)
        ttk.Button(buttons, text="清空", command=self.clear_files).pack(side="left")

        options = ttk.LabelFrame(frame, text="2. 排序设置", padding=12)
        options.pack(fill="x", pady=14)
        ttk.Label(options, text="排序列：").pack(side="left")
        self.column_box = ttk.Combobox(options, textvariable=self.sort_column, state="readonly", width=24)
        self.column_box.pack(side="left", padx=(4, 18))
        ttk.Radiobutton(options, text="升序", variable=self.descending, value=False).pack(side="left")
        ttk.Radiobutton(options, text="降序", variable=self.descending, value=True).pack(side="left", padx=12)

        ttk.Button(frame, text="3. 整理并另存为…", command=self.export).pack(anchor="w")
        ttk.Label(frame, textvariable=self.status, wraplength=650).pack(anchor="w", pady=(12, 0))

    def _refresh(self) -> None:
        self.file_list.delete(0, tk.END)
        for path in self.paths:
            self.file_list.insert(tk.END, str(path))
        if not self.paths:
            self.column_box["values"] = ()
            self.sort_column.set("")
            return
        try:
            headers = read_header(self.paths[0])
        except (OSError, ValueError, BadZipFile) as error:
            self.status.set(str(error))
            return
        self.column_box["values"] = headers
        if self.sort_column.get() not in headers:
            self.sort_column.set(headers[0])

    def add_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="选择 Excel 文件", filetypes=[("Excel 工作簿", "*.xlsx")]
        )
        for filename in selected:
            path = Path(filename).resolve()
            if path not in self.paths:
                self.paths.append(path)
        self.status.set(f"已选择 {len(self.paths)} 个文件。")
        self._refresh()

    def remove_selected(self) -> None:
        for index in reversed(self.file_list.curselection()):
            del self.paths[index]
        self._refresh()

    def clear_files(self) -> None:
        self.paths.clear()
        self._refresh()
        self.status.set("请选择要合并的 .xlsx 文件。")

    def export(self) -> None:
        if not self.paths:
            messagebox.showwarning("缺少文件", "请先添加至少一个 .xlsx 文件。")
            return
        destination = filedialog.asksaveasfilename(
            title="保存整理结果", defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")], initialfile="整理结果.xlsx",
        )
        if not destination:
            return
        try:
            result = merge_workbooks(self.paths, destination, self.sort_column.get(), self.descending.get())
        except (OSError, ValueError, TypeError, BadZipFile) as error:
            messagebox.showerror("整理失败", str(error))
            self.status.set("整理失败，请检查文件和表头。")
            return
        summary = (
            f"完成：读取 {result.input_rows} 行，删除 {result.duplicate_rows} 行重复数据，"
            f"输出 {result.output_rows} 行。\n保存位置：{result.output_path}"
        )
        self.status.set(summary)
        messagebox.showinfo("整理完成", summary)


if __name__ == "__main__":
    OrganizerApp().mainloop()
