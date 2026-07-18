from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from .errors import InvalidInputError, OperationCancelledError
from .excel import ExcelComparer, ExcelReader, ExcelReportWriter
from .models import CompareOptions, CompareResult

CancelCheck = Callable[[], bool]


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
        result = self.comparer.compare(self.reader.read(old), self.reader.read(new), options, cancel_requested)
        self._raise_if_cancelled(cancel_requested)
        self.writer.write(new, output, result, detailed, cancel_requested)
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
