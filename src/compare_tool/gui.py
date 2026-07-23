from __future__ import annotations

import os
import re
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
from .workbook_preparer import SUPPORTED_INPUT_EXTENSIONS

GUI_INPUT_EXTENSIONS = SUPPORTED_INPUT_EXTENSIONS | {".csv", ".json", ".xml"}
CSV_ENCODINGS = {
    "自動": "auto",
    "UTF-8 / UTF-8 BOM": "utf-8-sig",
    "Shift_JIS": "cp932",
}
CSV_DELIMITERS = {
    "自動": "auto",
    "カンマ ,": ",",
    "タブ": "\t",
    "セミコロン ;": ";",
}

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
        saved_options = self.settings.load_compare_options()
        saved_view_mode = self.settings.load_view_mode()
        self.last_output: Path | None = None
        self.is_busy = False
        self.cancel_event: threading.Event | None = None
        self.busy_controls: list[ttk.Widget] = []
        self.excel_options_frame: ttk.LabelFrame | None = None
        self.table_options_frame: ttk.LabelFrame | None = None
        self.csv_options_frame: ttk.LabelFrame | None = None
        self.json_options_frame: ttk.LabelFrame | None = None
        self.xml_options_frame: ttk.LabelFrame | None = None
        self.old_path = tk.StringVar()
        self.new_path = tk.StringVar()
        self.compare_values = tk.BooleanVar(value=saved_options.compare_values)
        self.compare_formulas = tk.BooleanVar(value=saved_options.compare_formulas)
        self.empty_equals_empty = tk.BooleanVar(value=saved_options.empty_string_equals_empty)
        self.ignore_whitespace = tk.BooleanVar(value=saved_options.ignore_surrounding_whitespace)
        self.ignore_case = tk.BooleanVar(value=saved_options.ignore_case)
        self.algorithm = tk.StringVar(value=saved_options.algorithm.value)
        self.key_columns = tk.StringVar(value=",".join(saved_options.key_columns))
        self.csv_encoding = tk.StringVar(value=self._csv_encoding_label(saved_options.csv_encoding))
        self.csv_delimiter = tk.StringVar(value=self._csv_delimiter_label(saved_options.csv_delimiter))
        self.ignore_csv_blank_lines = tk.BooleanVar(value=saved_options.ignore_csv_blank_lines)
        self.ignore_json_object_key_order = tk.BooleanVar(value=saved_options.ignore_json_object_key_order)
        self.ignore_json_array_order = tk.BooleanVar(value=saved_options.ignore_json_array_order)
        self.json_array_key = tk.StringVar(value=saved_options.json_array_key)
        self.ignore_xml_attribute_order = tk.BooleanVar(value=saved_options.ignore_xml_attribute_order)
        self.ignore_xml_blank_text = tk.BooleanVar(value=saved_options.ignore_xml_blank_text)
        self.view_mode = tk.StringVar(value=saved_view_mode)
        self.status = tk.StringVar(value="ファイルを指定してください")
        self._build()
        self.old_path.trace_add("write", self._on_input_path_changed)
        self.new_path.trace_add("write", self._on_input_path_changed)
        self._update_option_states()

    def _build(self) -> None:
        self.root.title(f"compare_tool v{__version__} - ファイル差分比較")
        self.root.geometry("920x650")
        self.root.minsize(760, 560)

        main = ttk.Frame(self.root, padding=14)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(5, weight=1)

        header = ttk.Frame(main)
        header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="ファイル差分比較", font=("Yu Gothic UI", 16, "bold")).grid(row=0, column=0, sticky="w")
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
        options.columnconfigure(0, weight=1)

        common_options = ttk.Frame(options)
        common_options.grid(row=0, column=0, sticky="ew")
        checks = [
            ("セル値を比較", self.compare_values),
            ("空文字と空セルを同一視", self.empty_equals_empty),
            ("前後スペースを無視", self.ignore_whitespace),
            ("大文字小文字を無視", self.ignore_case),
        ]
        for index, (text, variable) in enumerate(checks):
            checkbutton = ttk.Checkbutton(common_options, text=text, variable=variable)
            checkbutton.grid(row=index // 2, column=index % 2, sticky="w", padx=(0, 28), pady=3)
            self.busy_controls.append(checkbutton)
        ttk.Label(common_options, text="表示方法:").grid(row=2, column=0, sticky="w", pady=(9, 0))
        for column, text, value in [(1, "詳細表示", "detail"), (2, "サマリー表示", "summary")]:
            radio = ttk.Radiobutton(common_options, text=text, variable=self.view_mode, value=value)
            radio.grid(row=2, column=column, sticky="w", pady=(9, 0))
            self.busy_controls.append(radio)

        self.excel_options_frame = ttk.LabelFrame(options, text="Excelオプション", padding=8)
        self.excel_options_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        formula_check = ttk.Checkbutton(self.excel_options_frame, text="数式を比較", variable=self.compare_formulas)
        formula_check.grid(row=0, column=0, sticky="w")
        self.busy_controls.append(formula_check)

        self.table_options_frame = ttk.LabelFrame(options, text="表形式オプション", padding=8)
        self.table_options_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(self.table_options_frame, text="比較方式:").grid(row=0, column=0, sticky="w")
        algorithms = [
            ("セル座標比較", CompareAlgorithm.CELL_COORDINATE.value),
            ("行追加/削除を考慮", CompareAlgorithm.ROW_LCS.value),
            ("キー列で比較", CompareAlgorithm.KEY_COLUMNS.value),
        ]
        for column, (text, value) in enumerate(algorithms, 1):
            radio = ttk.Radiobutton(self.table_options_frame, text=text, variable=self.algorithm, value=value)
            radio.grid(row=0, column=column, sticky="w", padx=(0, 16))
            self.busy_controls.append(radio)
        ttk.Label(self.table_options_frame, text="キー列:").grid(row=1, column=0, sticky="w", pady=(9, 0))
        key_entry = ttk.Entry(self.table_options_frame, textvariable=self.key_columns, width=18)
        key_entry.grid(row=1, column=1, sticky="w", pady=(9, 0))
        self.busy_controls.append(key_entry)
        ttk.Label(self.table_options_frame, text="例: A または A,C").grid(
            row=1,
            column=2,
            sticky="w",
            pady=(9, 0),
        )

        self.csv_options_frame = ttk.LabelFrame(options, text="CSVオプション", padding=8)
        self.csv_options_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(self.csv_options_frame, text="文字コード:").grid(row=0, column=0, sticky="w")
        encoding_combo = ttk.Combobox(
            self.csv_options_frame,
            textvariable=self.csv_encoding,
            values=list(CSV_ENCODINGS),
            state="readonly",
            width=18,
        )
        encoding_combo.grid(row=0, column=1, sticky="w", padx=(8, 24))
        self.busy_controls.append(encoding_combo)
        ttk.Label(self.csv_options_frame, text="区切り文字:").grid(row=0, column=2, sticky="w")
        delimiter_combo = ttk.Combobox(
            self.csv_options_frame,
            textvariable=self.csv_delimiter,
            values=list(CSV_DELIMITERS),
            state="readonly",
            width=14,
        )
        delimiter_combo.grid(row=0, column=3, sticky="w", padx=(8, 0))
        self.busy_controls.append(delimiter_combo)
        blank_line_check = ttk.Checkbutton(
            self.csv_options_frame,
            text="空行を無視",
            variable=self.ignore_csv_blank_lines,
        )
        blank_line_check.grid(row=1, column=0, sticky="w", pady=(9, 0))
        self.busy_controls.append(blank_line_check)

        self.json_options_frame = ttk.LabelFrame(options, text="JSONオプション", padding=8)
        self.json_options_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            self.json_options_frame,
            text="JSON Path単位で比較します。配列は既定ではインデックス順で比較します。",
        ).grid(row=0, column=0, sticky="w")
        json_key_order_check = ttk.Checkbutton(
            self.json_options_frame,
            text="オブジェクトのキー順を無視",
            variable=self.ignore_json_object_key_order,
        )
        json_key_order_check.grid(row=1, column=0, sticky="w", pady=(9, 0))
        self.busy_controls.append(json_key_order_check)
        json_array_order_check = ttk.Checkbutton(
            self.json_options_frame,
            text="配列の順序を無視",
            variable=self.ignore_json_array_order,
        )
        json_array_order_check.grid(row=2, column=0, sticky="w", pady=(3, 0))
        self.busy_controls.append(json_array_order_check)
        ttk.Label(self.json_options_frame, text="配列キー:").grid(row=3, column=0, sticky="w", pady=(9, 0))
        json_array_key_entry = ttk.Entry(self.json_options_frame, textvariable=self.json_array_key, width=24)
        json_array_key_entry.grid(row=3, column=1, sticky="w", pady=(9, 0))
        self.busy_controls.append(json_array_key_entry)
        ttk.Label(self.json_options_frame, text="例: id または name").grid(
            row=3,
            column=2,
            sticky="w",
            padx=(8, 0),
            pady=(9, 0),
        )

        self.xml_options_frame = ttk.LabelFrame(options, text="XMLオプション", padding=8)
        self.xml_options_frame.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(
            self.xml_options_frame,
            text="XPath風パス単位で比較します。同名兄弟要素は出現順で比較します。",
        ).grid(row=0, column=0, sticky="w")
        xml_attribute_order_check = ttk.Checkbutton(
            self.xml_options_frame,
            text="属性順を無視",
            variable=self.ignore_xml_attribute_order,
        )
        xml_attribute_order_check.grid(row=1, column=0, sticky="w", pady=(9, 0))
        self.busy_controls.append(xml_attribute_order_check)
        xml_blank_text_check = ttk.Checkbutton(
            self.xml_options_frame,
            text="空白のみテキストを無視",
            variable=self.ignore_xml_blank_text,
        )
        xml_blank_text_check.grid(row=2, column=0, sticky="w", pady=(3, 0))
        self.busy_controls.append(xml_blank_text_check)

        log_frame = ttk.LabelFrame(main, text="ログ", padding=6)
        log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=13, wrap="word", state="disabled")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        actions = ttk.Frame(main)
        actions.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        actions.columnconfigure(1, weight=1)
        ttk.Label(actions, text="状態:").grid(row=0, column=0, sticky="w", padx=(0, 4))
        ttk.Label(actions, textvariable=self.status).grid(row=0, column=1, sticky="ew")
        self.open_button = ttk.Button(actions, text="出力ファイルを開く", command=self._open_output, state="disabled")
        self.open_button.grid(row=0, column=2, sticky="e", padx=(8, 0))
        self.open_folder_button = ttk.Button(
            actions,
            text="保存フォルダを開く",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_folder_button.grid(row=0, column=3, sticky="e", padx=(8, 0))
        self.cancel_button = ttk.Button(actions, text="キャンセル", command=self._cancel_compare, state="disabled")
        self.cancel_button.grid(row=0, column=4, sticky="e", padx=(8, 0))
        self.compare_button = ttk.Button(actions, text="比較開始", command=self._start_compare)
        self.compare_button.grid(row=0, column=5, sticky="e", padx=(8, 0))
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
        selected = filedialog.askopenfilename(
            title="比較ファイルを選択",
            filetypes=[
                ("対応ファイル", "*.xlsx *.xls *.csv *.json *.xml"),
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("JSON", "*.json"),
                ("XML", "*.xml"),
            ],
        )
        if selected:
            variable.set(selected)

    def _drop(self, event: Any, variable: tk.StringVar) -> None:
        if self.is_busy:
            return
        paths = self.root.tk.splitlist(event.data)
        if not paths:
            return
        path = Path(paths[0])
        if path.suffix.lower() not in GUI_INPUT_EXTENSIONS:
            messagebox.showwarning("対象外ファイル", ".xlsx、.xls、.csv、.json、.xml ファイルのみ指定できます。")
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
            if path.suffix.lower() not in GUI_INPUT_EXTENSIONS:
                messagebox.showwarning(
                    "対象外ファイル",
                    f"{label}は .xlsx、.xls、.csv、.json、.xml のいずれかを指定してください。",
                )
                self.status.set("対象外ファイル")
                self._log(f"{label}の対象外ファイルを拒否しました: {path}")
                return False
            if not path.is_file():
                messagebox.showwarning("ファイルが見つかりません", f"{label}が見つかりません。\n{path}")
                self.status.set("ファイルが見つかりません")
                self._log(f"{label}が見つかりません: {path}")
                return False
        if self._format_family(Path(old_value)) != self._format_family(Path(new_value)):
            messagebox.showwarning("形式不一致", "旧ファイルと新ファイルは同じ形式を指定してください。")
            self.status.set("形式不一致")
            self._log("旧ファイルと新ファイルの形式が異なるため比較を開始できません。")
            return False
        return self.algorithm.get() != CompareAlgorithm.KEY_COLUMNS.value or self._validate_key_columns()

    @staticmethod
    def _format_family(path: Path) -> str:
        suffix = path.suffix.lower()
        if suffix == ".csv":
            return "csv"
        if suffix == ".json":
            return "json"
        if suffix == ".xml":
            return "xml"
        return "excel"

    def _validate_key_columns(self) -> bool:
        columns = self._key_columns()
        if not columns:
            messagebox.showwarning("キー列未指定", "キー列で比較する場合はキー列を指定してください。\n例: A または A,C")
            self.status.set("キー列を指定してください")
            self._log("キー列未指定のため比較を開始できません。")
            return False
        invalid = [column for column in columns if not re.fullmatch(r"[A-Z]{1,3}", column)]
        if invalid:
            values = ", ".join(invalid)
            message = f"キー列は列名だけを指定してください。\n不正な指定: {values}\n例: A または A,C"
            messagebox.showwarning("キー列指定エラー", message)
            self.status.set("キー列指定エラー")
            self._log(f"キー列の不正な指定を拒否しました: {values}")
            return False
        if len(set(columns)) != len(columns):
            messagebox.showwarning("キー列指定エラー", "キー列が重複しています。")
            self.status.set("キー列指定エラー")
            self._log("重複したキー列指定を拒否しました。")
            return False
        return True

    def _on_input_path_changed(self, *_args: object) -> None:
        self._update_option_states()

    def _update_option_states(self) -> None:
        if self.is_busy:
            return

        family = self._selected_format_family()
        self._set_frame_visible(self.excel_options_frame, family == "excel")
        self._set_frame_visible(self.table_options_frame, family in {"excel", "csv"})
        self._set_frame_visible(self.csv_options_frame, family == "csv")
        self._set_frame_visible(self.json_options_frame, family == "json")
        self._set_frame_visible(self.xml_options_frame, family == "xml")

        if family == "csv":
            self.compare_formulas.set(False)
        elif family == "excel":
            return
        elif family in {"json", "xml"}:
            self.compare_formulas.set(False)
            if self.algorithm.get() != CompareAlgorithm.CELL_COORDINATE.value:
                self.algorithm.set(CompareAlgorithm.CELL_COORDINATE.value)
        else:
            self.compare_formulas.set(False)

    def _selected_format_family(self) -> str | None:
        old = self.old_path.get().strip()
        new = self.new_path.get().strip()
        if not old or not new:
            return None
        old_suffix = Path(old).suffix.lower()
        new_suffix = Path(new).suffix.lower()
        if old_suffix not in GUI_INPUT_EXTENSIONS or new_suffix not in GUI_INPUT_EXTENSIONS:
            return None
        old_family = self._format_family(Path(old))
        new_family = self._format_family(Path(new))
        if old_family != new_family:
            return None
        return old_family

    @staticmethod
    def _set_frame_visible(frame: ttk.LabelFrame | None, visible: bool) -> None:
        if frame is None:
            return
        if visible:
            frame.grid()
        else:
            frame.grid_remove()

    def _options(self) -> CompareOptions:
        return CompareOptions(
            compare_values=self.compare_values.get(),
            compare_formulas=self.compare_formulas.get(),
            empty_string_equals_empty=self.empty_equals_empty.get(),
            ignore_surrounding_whitespace=self.ignore_whitespace.get(),
            ignore_case=self.ignore_case.get(),
            algorithm=CompareAlgorithm(self.algorithm.get()),
            key_columns=self._key_columns(),
            csv_encoding=CSV_ENCODINGS[self.csv_encoding.get()],
            csv_delimiter=CSV_DELIMITERS[self.csv_delimiter.get()],
            ignore_csv_blank_lines=self.ignore_csv_blank_lines.get(),
            ignore_json_object_key_order=self.ignore_json_object_key_order.get(),
            ignore_json_array_order=self.ignore_json_array_order.get(),
            json_array_key=self.json_array_key.get().strip(),
            ignore_xml_attribute_order=self.ignore_xml_attribute_order.get(),
            ignore_xml_blank_text=self.ignore_xml_blank_text.get(),
        )

    @staticmethod
    def _csv_encoding_label(value: str) -> str:
        for label, encoding in CSV_ENCODINGS.items():
            if encoding == value:
                return label
        return "自動"

    @staticmethod
    def _csv_delimiter_label(value: str) -> str:
        for label, delimiter in CSV_DELIMITERS.items():
            if delimiter == value:
                return label
        return "自動"

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
        self.settings.save_compare_options(self._options())
        self.settings.save_view_mode(self.view_mode.get())
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
        if not busy:
            self._update_option_states()

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
