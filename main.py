"""Tkinter desktop entry point for Excel日报自动处理助手."""
from __future__ import annotations

import queue
import threading
import traceback
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from processor import ProcessingError, process_many
from excel_clean_split import CleanSplitWindow
from xls_clean_split import XlsCleanSplitWindow


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel日报自动处理助手")
        self.geometry("820x580")
        self.minsize(720, 500)
        self.brand_files: list[Path] = []
        self.import_file: Path | None = None
        self.output_dir: Path | None = None
        self.events: queue.Queue[tuple[str, str]] = queue.Queue()
        self._build_ui()
        self.after(100, self._drain_events)

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=16)
        container.pack(fill="both", expand=True)

        ttk.Button(
            container,
            text="Excel数据清理分列",
            command=self.open_clean_split,
        ).pack(anchor="e", pady=(0, 10))

        ttk.Button(
            container,
            text="XLS数据清理转换",
            command=self.open_xls_clean_split,
        ).pack(anchor="e", pady=(0, 10))

        ttk.Label(container, text="Excel日报自动处理助手", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(container, text="批量处理品牌日报，保留模板格式与公式结构").pack(anchor="w", pady=(2, 14))

        controls = ttk.LabelFrame(container, text="文件选择", padding=12)
        controls.pack(fill="x")
        ttk.Label(controls, text="处理日期").grid(row=0, column=0, sticky="w", pady=4)
        self.date_entry = ttk.Entry(controls, width=20)
        today = date.today()
        self.date_entry.insert(0, f"{today.year}/{today.month}/{today.day}")
        self.date_entry.grid(row=0, column=1, sticky="w", padx=12)
        ttk.Button(controls, text="1. 选择多个品牌日报 XLSX", command=self.select_brands).grid(row=1, column=0, sticky="ew", pady=4)
        self.brand_label = ttk.Label(controls, text="尚未选择", foreground="#555")
        self.brand_label.grid(row=1, column=1, sticky="w", padx=12)
        ttk.Button(controls, text="2. 选择销售数据 XLSX", command=self.select_import).grid(row=2, column=0, sticky="ew", pady=4)
        self.import_label = ttk.Label(controls, text="尚未选择", foreground="#555")
        self.import_label.grid(row=2, column=1, sticky="w", padx=12)
        ttk.Button(controls, text="3. 选择输出目录", command=self.select_output).grid(row=3, column=0, sticky="ew", pady=4)
        self.output_label = ttk.Label(controls, text="尚未选择", foreground="#555")
        self.output_label.grid(row=3, column=1, sticky="w", padx=12)
        controls.columnconfigure(1, weight=1)

        action = ttk.Frame(container)
        action.pack(fill="x", pady=12)
        self.start_button = ttk.Button(action, text="开始处理", command=self.start)
        self.start_button.pack(side="left")
        self.progress = ttk.Progressbar(action, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(12, 0))

        ttk.Label(container, text="处理日志").pack(anchor="w")
        self.log = tk.Text(container, height=18, state="disabled", wrap="word", font=("Consolas", 10))
        self.log.pack(fill="both", expand=True, pady=(4, 0))

    def open_clean_split(self) -> None:
        """打开与日报处理逻辑隔离的 Excel 数据清理页面。"""
        CleanSplitWindow(self)

    def open_xls_clean_split(self) -> None:
        """打开独立的 XLS 转换清理页面。"""
        XlsCleanSplitWindow(self)

    def select_brands(self) -> None:
        paths = filedialog.askopenfilenames(title="选择品牌日报", filetypes=[("Excel 工作簿", "*.xlsx")])
        if paths:
            self.brand_files = [Path(path) for path in paths]
            self.brand_label.config(text=f"已选择 {len(paths)} 个文件")

    def select_import(self) -> None:
        path = filedialog.askopenfilename(title="选择销售数据文件", filetypes=[("Excel 工作簿", "*.xlsx")])
        if path:
            self.import_file = Path(path)
            self.import_label.config(text=self.import_file.name)

    def select_output(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir = Path(path)
            self.output_label.config(text=str(self.output_dir))

    def start(self) -> None:
        if not self.brand_files:
            messagebox.showwarning("缺少文件", "请选择至少一个品牌日报 XLSX 文件。")
            return
        if not self.import_file:
            messagebox.showwarning("缺少文件", "请选择销售数据 XLSX 文件。")
            return
        if not self.output_dir:
            messagebox.showwarning("缺少目录", "请选择输出目录。")
            return
        try:
            self.processing_date = self._parse_date(self.date_entry.get().strip())
        except ValueError:
            messagebox.showwarning("日期错误", "处理日期格式应为 YYYY/M/D，例如 2026/8/1。")
            return
        self.start_button.config(state="disabled")
        self.progress.start(10)
        self._append_log("开始处理，请勿打开或修改相关文件。")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self) -> None:
        try:
            outputs = process_many(self.brand_files, self.import_file, self.output_dir, self.processing_date,
                                   lambda text: self.events.put(("log", text)))
            self.events.put(("done", f"处理完成，共输出 {len(outputs)} 个文件。"))
        except ProcessingError as exc:
            self.events.put(("error", str(exc)))
        except Exception as exc:
            traceback.print_exc()
            self.events.put(("error", f"未预期错误：{exc}"))

    def _drain_events(self) -> None:
        try:
            while True:
                kind, text = self.events.get_nowait()
                if kind == "log":
                    self._append_log(text)
                elif kind == "done":
                    self._finish()
                    self._append_log(text)
                    messagebox.showinfo("完成", text)
                elif kind == "error":
                    self._finish()
                    self._append_log("错误：" + text)
                    messagebox.showerror("处理失败", text)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)

    def _finish(self) -> None:
        self.progress.stop()
        self.start_button.config(state="normal")

    @staticmethod
    def _parse_date(text: str) -> date:
        for pattern in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, pattern).date()
            except ValueError:
                continue
        raise ValueError(text)

    def _append_log(self, text: str) -> None:
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    Application().mainloop()
