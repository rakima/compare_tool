from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from .errors import InvalidInputError, OperationCancelledError
from .excel import ExcelComparer, ExcelReader, ExcelReportWriter
from .models import CompareOptions, CompareResult

CancelCheck = Callable[[], bool]
ProgressCallback = Callable[[str], None]
LARGE_DIFFERENCE_NOTICE_THRESHOLD = 1_000


class CompareUseCase:
    def __init__(
        self,
        reader: ExcelReader | None = None,
        comparer: ExcelComparer | None = None,
        writer: ExcelReportWriter | None = None,
    ) -> None:
        self.reader = reader or ExcelReader()
        self.comparer = comparer or ExcelComparer()
        self.writer = writer or ExcelReportWriter()

    def execute(
        self,
        old_path: str | Path,
        new_path: str | Path,
        output_path: str | Path,
        options: CompareOptions,
        detailed: bool = True,
        cancel_requested: CancelCheck | None = None,
        progress_callback: ProgressCallback | None = None,
    ) -> CompareResult:
        old = self._validate(old_path, "旧ファイル")
        new = self._validate(new_path, "新ファイル")
        output = Path(output_path).expanduser()
        if self._same_file(old, new):
            raise InvalidInputError("旧ファイルと新ファイルに同じファイルが指定されています。")
        if output.suffix.lower() != ".xlsx":
            raise InvalidInputError("出力ファイルの拡張子は .xlsx にしてください。")
        if self._same_file(new, output) or self._same_file(old, output):
            raise InvalidInputError("出力先には入力ファイルと異なるパスを指定してください。")

        self._raise_if_cancelled(cancel_requested)
        self._notify(progress_callback, "旧ファイルを読み込んでいます...")
        old_document = self.reader.read(old)
        self._raise_if_cancelled(cancel_requested)
        self._notify(progress_callback, "新ファイルを読み込んでいます...")
        new_document = self.reader.read(new)
        self._raise_if_cancelled(cancel_requested)
        self._notify(progress_callback, "差分を検出しています...")
        result = self.comparer.compare(old_document, new_document, options, cancel_requested)
        self._raise_if_cancelled(cancel_requested)
        self._notify(progress_callback, f"差分を {result.total:,} 件検出しました。")
        if detailed and result.total >= LARGE_DIFFERENCE_NOTICE_THRESHOLD:
            self._notify(progress_callback, "差分が多いため、詳細レポートの作成に時間がかかる場合があります。")
        self._notify(progress_callback, "比較結果Excelを作成しています...")
        self.writer.write(new, output, result, detailed, cancel_requested)
        self._notify(progress_callback, "比較結果Excelの作成が完了しました。")
        return result

    @staticmethod
    def _validate(path_value: str | Path, label: str) -> Path:
        path = Path(path_value).expanduser()
        if path.suffix.lower() != ".xlsx":
            raise InvalidInputError(f"{label}は .xlsx ファイルを指定してください。")
        if not path.is_file():
            raise InvalidInputError(f"{label}が見つかりません: {path}")
        return path

    @staticmethod
    def _same_file(left: Path, right: Path) -> bool:
        try:
            return left.resolve().samefile(right.resolve())
        except (FileNotFoundError, OSError):
            return str(left.resolve()).casefold() == str(right.resolve()).casefold()

    @staticmethod
    def _raise_if_cancelled(cancel_requested: CancelCheck | None) -> None:
        if cancel_requested is not None and cancel_requested():
            raise OperationCancelledError("比較をキャンセルしました。")

    @staticmethod
    def _notify(progress_callback: ProgressCallback | None, message: str) -> None:
        if progress_callback is not None:
            progress_callback(message)
