from __future__ import annotations

import os
import tempfile
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import WorkbookConversionError

SUPPORTED_INPUT_EXTENSIONS = frozenset({".xlsx", ".xls"})
PASS_THROUGH_EXTENSIONS = frozenset({".xlsx"})


@dataclass(frozen=True)
class PreparedWorkbook:
    """Workbook path that can be consumed by the Excel comparison pipeline."""

    source_path: Path
    prepared_path: Path
    converted: bool = False


class ExcelWorkbookConverter:
    """Converts legacy `.xls` workbooks to `.xlsx` by automating Excel on Windows."""

    XLSX_FILE_FORMAT = 51

    def convert(self, path: Path) -> Path:
        output = self._temporary_xlsx_path(path)
        excel: Any | None = None
        workbook: Any | None = None
        try:
            try:
                import win32com.client
            except ImportError as exc:
                raise WorkbookConversionError(
                    ".xls ファイルを変換するには、Windows版Excelとpywin32が必要です。"
                    "Excelをインストールした環境で依存関係を入れ直してから再実行してください。"
                ) from exc

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(
                str(path.resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                Password="",
            )
            workbook.SaveAs(str(output), FileFormat=self.XLSX_FILE_FORMAT)
            return output
        except WorkbookConversionError:
            with suppress(OSError):
                output.unlink(missing_ok=True)
            raise
        except Exception as exc:
            with suppress(OSError):
                output.unlink(missing_ok=True)
            raise WorkbookConversionError(
                f".xls ファイルをExcelで .xlsx に変換できませんでした: {path}\n"
                "ファイルが破損していないか、パスワード付きではないか、Excelで開けるか確認してください。"
            ) from exc
        finally:
            if workbook is not None:
                with suppress(Exception):
                    workbook.Close(False)
            if excel is not None:
                with suppress(Exception):
                    excel.Quit()

    @staticmethod
    def _temporary_xlsx_path(source: Path) -> Path:
        file_descriptor, temporary_name = tempfile.mkstemp(prefix=f"{source.stem}_", suffix=".converted.xlsx")
        os.close(file_descriptor)
        return Path(temporary_name)


class WorkbookPreparer:
    """Prepare input workbooks before they are read by openpyxl.

    `.xlsx` files pass through unchanged. Legacy `.xls` workbooks are converted
    to temporary `.xlsx` files before openpyxl reads them.
    """

    def __init__(self, converter: ExcelWorkbookConverter | None = None) -> None:
        self.converter = converter or ExcelWorkbookConverter()

    def prepare(self, path: Path) -> PreparedWorkbook:
        suffix = path.suffix.lower()
        if suffix in PASS_THROUGH_EXTENSIONS:
            return PreparedWorkbook(source_path=path, prepared_path=path)
        if suffix == ".xls":
            converted = self.converter.convert(path)
            return PreparedWorkbook(source_path=path, prepared_path=converted, converted=True)
        raise WorkbookConversionError(f"対応していないExcel形式です: {path.suffix}")
