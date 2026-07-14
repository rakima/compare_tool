from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .errors import CompareToolError
from .models import CompareOptions
from .usecase import CompareUseCase

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # The application remains usable without optional DnD support.
    DND_FILES = None
    TkinterDnD = None


class CompareApp:
    def __init__(self, root: tk.Tk, use_case: CompareUseCase | None = None) -> None:
        self.root = root
        self.use_case = use_case or CompareUseCase()
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.compare_values = tk.BooleanVar(value=True)
        self.compare_formulas = tk.BooleanVar(value=True)
        self.empty_equals_empty = tk.BooleanVar(value=True)
        self.ignore_whitespace = tk.BooleanVar(value=False)
        self.ignore_case = tk.BooleanVar(value=False)
        self.view_mode = tk.StringVar(value="detail")
        self.status = tk.StringVar(value="ファイルを指定してください")
        self._build()

    def _build(self) -> None:
        self.root.title("compare_tool - Excel差分比較")
        self.root.geometry("920x650")
        self.root.minsize(760, 560)

        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(5, weight=1)

        ttk.Label(main, text="Excel差分比較", font=("Yu Gothic UI", 16, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 12)
        )
        self._file_row(main, 1, "旧ファイル", self.old_path)
        self._file_row(main, 2, "新ファイル", self.new_path)

        options = ttk.LabelFrame(main, text="比較オプション", padding=10)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=12)
        checks = [
            ("セル値を比較", self.compare_values),
            ("数式を比較", self.compare_formulas),
            ("空文字と空セルを同一視", self.empty_equals_empty),
            ("前後スペースを無視", self.ignore_whitespace),
            ("大文字小文字を無視", self.ignore_case),
        ]
        for index, (text, variable) in enumerate(checks):
            ttk.Checkbutton(options, text=text, variable=variable).grid(
                row=index // 3, column=index % 3, sticky="w", padx=(0, 28), pady=3
            )
        ttk.Label(options, text="表示方法:").grid(row=2, column=0, sticky="w", pady=(9, 0))
        ttk.Radiobutton(options, text="詳細表示", variable=self.view_mode, value="detail").grid(
            row=2, column=1, sticky="w", pady=(9, 0)
        )
        ttk.Radiobutton(options, text="サマリー表示", variable=self.view_mode, value="summary").grid(
            row=2, column=2, sticky="w", pady=(9, 0)
        )

        actions = ttk.Frame(main)
        actions.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        self.compare_button = ttk.Button(actions, text="比較開始", command=self._start_compare)
        self.compare_button.pack(side="left")
        ttk.Label(actions, textvariable=self.status).pack(side="left", padx=12)

        log_frame = ttk.LabelFrame(main, text="ログ", padding=6)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=13, wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self._log("起動しました。旧ファイルと新ファイルを指定してください。")

    def _file_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar) -> None:
        ttk.Label(parent, text=label, width=10).grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=5)
        ttk.Button(parent, text="ファイル選択...", command=lambda: self._browse(variable)).grid(
            row=row, column=2, pady=5
        )
        if DND_FILES is not None and hasattr(entry, "drop_target_register"):
            entry.drop_target_register(DND_FILES)
            entry.dnd_bind("<<Drop>>", lambda event: self._drop(event, variable))

    def _browse(self, variable: tk.StringVar) -> None:
        selected = filedialog.askopenfilename(title="Excelファイルを選択", filetypes=[("Excel", "*.xlsx")])
        if selected:
            variable.set(selected)

    def _drop(self, event, variable: tk.StringVar) -> None:
        paths = self.root.tk.splitlist(event.data)
        if paths:
            variable.set(paths[0])

    def _start_compare(self) -> None:
        if not self.compare_values.get() and not self.compare_formulas.get():
            messagebox.showwarning("比較オプション", "セル値または数式の少なくとも一方を選択してください。")
            return
        new_path = Path(self.new_path.get()) if self.new_path.get() else Path("新ファイル.xlsx")
        suggested = f"{new_path.stem}_比較結果.xlsx"
        output = filedialog.asksaveasfilename(
            title="比較結果の保存先", defaultextension=".xlsx", initialfile=suggested,
            filetypes=[("Excel", "*.xlsx")]
        )
        if not output:
            return
        self.compare_button.configure(state="disabled")
        self.status.set("比較中...")
        self._log("比較を開始します。")
        args = (self.old_path.get(), self.new_path.get(), output, self._options(), self.view_mode.get() == "detail")
        threading.Thread(target=self._run_compare, args=args, daemon=True).start()

    def _options(self) -> CompareOptions:
        return CompareOptions(
            compare_values=self.compare_values.get(), compare_formulas=self.compare_formulas.get(),
            empty_string_equals_empty=self.empty_equals_empty.get(),
            ignore_surrounding_whitespace=self.ignore_whitespace.get(), ignore_case=self.ignore_case.get(),
        )

    def _run_compare(self, old: str, new: str, output: str, options: CompareOptions, detailed: bool) -> None:
        try:
            result = self.use_case.execute(old, new, output, options, detailed)
            summary = result.summary()
            detail = "、".join(f"{kind.value}: {count}" for kind, count in summary.items())
            self.root.after(0, lambda: self._success(output, detail))
        except CompareToolError as exc:
            self.root.after(0, lambda exc=exc: self._failure(str(exc)))
        except Exception as exc:
            self.root.after(0, lambda exc=exc: self._failure(f"予期しないエラーが発生しました: {exc}"))

    def _success(self, output: str, detail: str) -> None:
        self.compare_button.configure(state="normal")
        self.status.set("比較完了")
        self._log(f"比較完了: {detail}\n出力先: {output}")
        messagebox.showinfo("比較完了", f"比較結果を保存しました。\n{output}")

    def _failure(self, message: str) -> None:
        self.compare_button.configure(state="normal")
        self.status.set("エラー")
        self._log(f"エラー: {message}")
        messagebox.showerror("エラー", message)

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{timestamp}] {message}\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main() -> None:
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
    CompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

