from __future__ import annotations

import os
import threading
import tkinter as tk
from datetime import datetime
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, cast

from . import __version__
from .errors import CompareToolError, OperationCancelledError
from .models import CompareAlgorithm, CompareOptions
from .settings import AppSettingsStore
from .usecase import CompareUseCase

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # The application remains usable without optional DnD support.
    DND_FILES = None
    TkinterDnD = None


class CompareApp:
    def __init__(
        self,
        root: tk.Tk,
        use_case: CompareUseCase | None = None,
        settings: AppSettingsStore | None = None,
    ) -> None:
        self.root = root
        self.use_case = use_case or CompareUseCase()
        self.settings = settings or AppSettingsStore()
        self.last_save_dir = self.settings.load_last_save_dir()
        self.file_history = self.settings.load_file_history()
        self.last_output: Path | None = None
        self.is_busy = False
        self.cancel_event: threading.Event | None = None
        self.busy_controls: list[ttk.Widget] = []
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.compare_values = tk.BooleanVar(value=True)
        self.compare_formulas = tk.BooleanVar(value=True)
        self.empty_equals_empty = tk.BooleanVar(value=True)
        self.ignore_whitespace = tk.BooleanVar(value=False)
        self.ignore_case = tk.BooleanVar(value=False)
        self.algorithm = tk.StringVar(value=CompareAlgorithm.CELL_COORDINATE.value)
        self.key_columns = tk.StringVar()
        self.view_mode = tk.StringVar(value="detail")
        self.status = tk.StringVar(value="ファイルを指定してください")
        self._build()

    def _build(self) -> None:
        self.root.title(f"compare_tool v{__version__} - Excel差分比較")
        self.root.geometry("920x650")
        self.root.minsize(760, 560)

        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(6, weight=1)

        header = ttk.Frame(main)
        header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Excel差分比較", font=("Yu Gothic UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=f"v{__version__}").grid(row=0, column=1, sticky="e")
        self.old_entry = self._file_row(main, 1, "旧ファイル", self.old_path)
        self.new_entry = self._file_row(main, 2, "新ファイル", self.new_path)
        file_actions = ttk.Frame(main)
        file_actions.grid(row=3, column=1, columnspan=2, sticky="e", pady=(2, 0))
        self.swap_button = ttk.Button(file_actions, text="旧/新を入替", command=self._swap_inputs)
        self.swap_button.pack(side="left")
        self.clear_history_button = ttk.Button(file_actions, text="履歴クリア", command=self._clear_history)
        self.clear_history_button.pack(side="left", padx=(8, 0))
        self.busy_controls.extend([self.swap_button, self.clear_history_button])

        options = ttk.LabelFrame(main, text="比較オプション", padding=10)
        options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=12)
        checks = [
            ("セル値を比較", self.compare_values),
            ("数式を比較", self.compare_formulas),
            ("空文字と空セルを同一視", self.empty_equals_empty),
            ("前後スペースを無視", self.ignore_whitespace),
            ("大文字小文字を無視", self.ignore_case),
        ]
        for index, (text, variable) in enumerate(checks):
            checkbutton = ttk.Checkbutton(options, text=text, variable=variable)
            checkbutton.grid(row=index // 3, column=index % 3, sticky="w", padx=(0, 28), pady=3)
            self.busy_controls.append(checkbutton)
        ttk.Label(options, text="表示方法:").grid(row=2, column=0, sticky="w", pady=(9, 0))
        for column, text, value in [(1, "詳細表示", "detail"), (2, "サマリー表示", "summary")]:
            radio = ttk.Radiobutton(options, text=text, variable=self.view_mode, value=value)
            radio.grid(row=2, column=column, sticky="w", pady=(9, 0))
            self.busy_controls.append(radio)
        ttk.Label(options, text="比較方式:").grid(row=3, column=0, sticky="w", pady=(9, 0))
        algorithms = [
            ("セル座標比較", CompareAlgorithm.CELL_COORDINATE.value),
            ("行追加/削除を考慮", CompareAlgorithm.ROW_LCS.value),
            ("キー列で比較", CompareAlgorithm.KEY_COLUMNS.value),
        ]
        for column, (text, value) in enumerate(algorithms, 1):
            radio = ttk.Radiobutton(options, text=text, variable=self.algorithm, value=value)
            radio.grid(row=3, column=column, sticky="w", pady=(9, 0))
            self.busy_controls.append(radio)
        ttk.Label(options, text="キー列:").grid(row=4, column=0, sticky="w", pady=(9, 0))
        key_entry = ttk.Entry(options, textvariable=self.key_columns, width=18)
        key_entry.grid(row=4, column=1, sticky="w", pady=(9, 0))
        self.busy_controls.append(key_entry)

        actions = ttk.Frame(main)
        actions.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        self.compare_button = ttk.Button(actions, text="比較開始", command=self._start_compare)
        self.compare_button.pack(side="left")
        self.cancel_button = ttk.Button(actions, text="キャンセル", command=self._cancel_compare, state="disabled")
        self.cancel_button.pack(side="left", padx=(8, 0))
        self.open_button = ttk.Button(actions, text="出力ファイルを開く", command=self._open_output, state="disabled")
        self.open_button.pack(side="left", padx=(8, 0))
        self.open_folder_button = ttk.Button(
            actions,
            text="保存フォルダを開く",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_button.pack(side="left", padx=(8, 0))
        ttk.Label(actions, textvariable=self.status).pack(side="left", padx=12)
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=130)
        self.progress.pack(side="right")

        log_frame = ttk.LabelFrame(main, text="ログ", padding=6)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=13, wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self._log("起動しました。旧ファイルと新ファイルを指定してください。")

    def _file_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar) -> ttk.Combobox:
        ttk.Label(parent, text=label, width=10).grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Combobox(parent, textvariable=variable, values=self.file_history)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=5)
        browse_button = ttk.Button(parent, text="ファイル選択...", command=lambda: self._browse(variable))
        browse_button.grid(row=row, column=2, pady=5)
        self.busy_controls.extend([entry, browse_button])
        if DND_FILES is not None and hasattr(entry, "drop_target_register"):
            entry.drop_target_register(DND_FILES)
            cast(Any, entry).dnd_bind("<<Drop>>", lambda event: self._drop(event, variable))
        return entry

    def _browse(self, variable: tk.StringVar) -> None:
        selected = filedialog.askopenfilename(title="Excelファイルを選択", filetypes=[("Excel", "*.xlsx")])
        if selected:
            variable.set(selected)

    def _drop(self, event: Any, variable: tk.StringVar) -> None:
        if self.is_busy:
            return
        paths = self.root.tk.splitlist(event.data)
        if not paths:
            return
        path = Path(paths[0])
        if path.suffix.lower() != ".xlsx":
            messagebox.showwarning("対象外ファイル", ".xlsx ファイルのみ指定できます。")
            self._log(f"対象外ファイルを拒否しました: {path}")
            return
        variable.set(str(path))

    def _start_compare(self) -> None:
        if not self._validate_inputs():
            return
        if not self.compare_values.get() and not self.compare_formulas.get():
            messagebox.showwarning("比較オプション", "セル値または数式の少なくとも一方を選択してください。")
            return
        new_path = Path(self.new_path.get())
        suggested = f"{new_path.stem}_比較結果.xlsx"
        initial_dir = self._default_save_dir(new_path)
        self._log(f"保存先候補: {initial_dir / suggested}")
        output = filedialog.asksaveasfilename(
            title="比較結果の保存先",
            defaultextension=".xlsx",
            initialfile=suggested,
            initialdir=str(initial_dir),
            filetypes=[("Excel", "*.xlsx")],
        )
        if not output:
            return
        self.last_output = None
        self.open_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self._set_busy(True)
        self.cancel_event = threading.Event()
        self.status.set("比較中...")
        self._log("比較を開始します。")
        if self.view_mode.get() == "detail":
            self._log("詳細表示では、差分が多い場合にレポート作成へ時間がかかることがあります。")
        args = (self.old_path.get(), self.new_path.get(), output, self._options(), self.view_mode.get() == "detail")
        threading.Thread(target=self._run_compare, args=args, daemon=True).start()

    def _validate_inputs(self) -> bool:
        old_value = self.old_path.get().strip()
        new_value = self.new_path.get().strip()
        if not old_value or not new_value:
            messagebox.showwarning("ファイル未指定", "旧ファイルと新ファイルを指定してください。")
            self.status.set("ファイルを指定してください")
            self._log("ファイル未指定のため比較を開始できません。")
            return False

        for label, value in [("旧ファイル", old_value), ("新ファイル", new_value)]:
            path = Path(value)
            if path.suffix.lower() != ".xlsx":
                messagebox.showwarning("対象外ファイル", f"{label}は .xlsx ファイルを指定してください。")
                self.status.set("対象外ファイル")
                self._log(f"{label}の対象外ファイルを拒否しました: {path}")
                return False
            if not path.is_file():
                messagebox.showwarning("ファイルが見つかりません", f"{label}が見つかりません。\n{path}")
                self.status.set("ファイルが見つかりません")
                self._log(f"{label}が見つかりません: {path}")
                return False
        return True

    def _options(self) -> CompareOptions:
        return CompareOptions(
            compare_values=self.compare_values.get(),
            compare_formulas=self.compare_formulas.get(),
            empty_string_equals_empty=self.empty_equals_empty.get(),
            ignore_surrounding_whitespace=self.ignore_whitespace.get(),
            ignore_case=self.ignore_case.get(),
            algorithm=CompareAlgorithm(self.algorithm.get()),
            key_columns=self._key_columns(),
        )

    def _key_columns(self) -> tuple[str, ...]:
        columns: list[str] = []
        for value in self.key_columns.get().replace("，", ",").split(","):
            column = value.strip().upper()
            if column:
                columns.append(column)
        return tuple(columns)

    def _run_compare(self, old: str, new: str, output: str, options: CompareOptions, detailed: bool) -> None:
        try:
            cancel_requested = self.cancel_event.is_set if self.cancel_event is not None else None
            result = self.use_case.execute(
                old,
                new,
                output,
                options,
                detailed,
                cancel_requested,
                self._queue_progress,
            )
            summary = result.summary()
            detail = "、".join(f"{kind.value}: {count}" for kind, count in summary.items())
            self.root.after(0, partial(self._success, output, detail))
        except OperationCancelledError:
            self.root.after(0, self._cancelled)
        except CompareToolError as exc:
            self.root.after(0, partial(self._failure, str(exc)))
        except Exception as exc:
            self.root.after(0, partial(self._failure, f"予期しないエラーが発生しました: {exc}"))

    def _queue_progress(self, message: str) -> None:
        self.root.after(0, partial(self._progress, message))

    def _progress(self, message: str) -> None:
        self.status.set(message)
        self._log(message)

    def _success(self, output: str, detail: str) -> None:
        self.last_save_dir = Path(output).resolve().parent
        self.settings.save_last_save_dir(self.last_save_dir)
        self._remember_inputs()
        self.last_output = Path(output).resolve()
        self._set_busy(False)
        self.cancel_event = None
        self.open_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        self.status.set("比較完了")
        self._log(f"比較完了: {detail}\n出力先: {output}")
        if messagebox.askyesno("比較完了", f"比較結果を保存しました。\n{output}\n\n出力ファイルを開きますか？"):
            self._open_output()

    def _failure(self, message: str) -> None:
        self._set_busy(False)
        self.cancel_event = None
        self.open_folder_button.configure(state="disabled")
        self.status.set("エラー")
        self._log(f"エラー: {message}")
        messagebox.showerror("エラー", message)

    def _cancelled(self) -> None:
        self._set_busy(False)
        self.cancel_event = None
        self.open_folder_button.configure(state="disabled")
        self.status.set("キャンセルしました")
        self._log("比較をキャンセルしました。出力ファイルは作成していません。")

    def _cancel_compare(self) -> None:
        if self.cancel_event is None or self.cancel_event.is_set():
            return
        self.cancel_event.set()
        self.cancel_button.configure(state="disabled")
        self.status.set("キャンセル中...")
        self._log("キャンセルを要求しました。処理中のステップが終わり次第停止します。")

    def _set_busy(self, busy: bool) -> None:
        self.is_busy = busy
        state = "disabled" if busy else "normal"
        self.compare_button.configure(state=state)
        self.cancel_button.configure(state="normal" if busy else "disabled")
        for control in self.busy_controls:
            control.state(["disabled"] if busy else ["!disabled"])
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _remember_inputs(self) -> None:
        current = [self.old_path.get(), self.new_path.get()]
        self.file_history = current + [path for path in self.file_history if path not in current]
        self.settings.save_file_history(self.file_history)
        self._refresh_history()

    def _refresh_history(self) -> None:
        self.file_history = self.settings.load_file_history()
        self.old_entry.configure(values=self.file_history)
        self.new_entry.configure(values=self.file_history)

    def _swap_inputs(self) -> None:
        old_value = self.old_path.get()
        self.old_path.set(self.new_path.get())
        self.new_path.set(old_value)
        self._log("旧ファイルと新ファイルを入れ替えました。")

    def _clear_history(self) -> None:
        if not self.file_history:
            self._log("ファイル履歴は空です。")
            return
        self.settings.clear_file_history()
        self._refresh_history()
        self._log("ファイル履歴をクリアしました。")

    def _open_output(self) -> None:
        output = self._validated_last_output()
        if output is None:
            return
        try:
            os.startfile(output)
        except OSError as exc:
            self._log(f"出力ファイルを開けません: {exc}")
            messagebox.showerror("ファイルを開けません", str(exc))

    def _open_output_folder(self) -> None:
        output = self._validated_last_output()
        if output is None:
            return
        try:
            os.startfile(output.parent)
        except OSError as exc:
            self._log(f"保存フォルダを開けません: {exc}")
            messagebox.showerror("フォルダを開けません", str(exc))

    def _validated_last_output(self) -> Path | None:
        if self.last_output is None or not self.last_output.is_file():
            messagebox.showwarning("ファイルを開けません", "出力ファイルが見つかりません。")
            self.open_button.configure(state="disabled")
            self.open_folder_button.configure(state="disabled")
            return None
        return self.last_output

    def _default_save_dir(self, new_path: Path) -> Path:
        if self.last_save_dir is not None and self.last_save_dir.is_dir():
            return self.last_save_dir
        new_directory = new_path.expanduser().resolve().parent
        return new_directory if new_directory.is_dir() else Path.cwd()

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
