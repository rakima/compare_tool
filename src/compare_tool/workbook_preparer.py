from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .errors import WorkbookConversionError

SUPPORTED_INPUT_EXTENSIONS = frozenset({".xlsx", ".xls"})
PASS_THROUGH_EXTENSIONS = frozenset({".xlsx"})


@dataclass(frozen=True)
class PreparedWorkbook:
    """Workbook path that can be consumed by the Excel comparison pipeline."""

    source_path: Path
    prepared_path: Path
    converted: bool = False


class WorkbookPreparer:
    """Prepare input workbooks before they are read by openpyxl.

    The current implementation passes `.xlsx` through unchanged and reserves a
    clear extension point for converting legacy `.xls` workbooks to temporary
    `.xlsx` files. A future converter can subclass or replace this class without
    changing the use case, comparer, or GUI code.
    """

    def prepare(self, path: Path) -> PreparedWorkbook:
        suffix = path.suffix.lower()
        if suffix in PASS_THROUGH_EXTENSIONS:
            return PreparedWorkbook(source_path=path, prepared_path=path)
        if suffix == ".xls":
            raise WorkbookConversionError(
                ".xls ファイルの変換機能はまだ実装されていません。"
                "将来のバージョンでExcelまたはLibreOfficeを使った変換に対応予定です。"
            )
        raise WorkbookConversionError(f"対応していないExcel形式です: {path.suffix}")
