from __future__ import annotations

import csv
import os
import tempfile
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

from .comparer import CancelCheck, Comparer
from .errors import OutputWriteError, WorkbookReadError
from .excel import CellData, ExcelComparer, ExcelDocument, ExcelReportWriter, raise_if_cancelled
from .models import CompareOptions, CompareResult

CSV_SHEET_NAME = "CSV"


@dataclass(slots=True)
class CsvDocument:
    rows: list[list[str]]

    @property
    def cells(self) -> dict[str, CellData]:
        cells: dict[str, CellData] = {}
        for row_index, row in enumerate(self.rows, 1):
            for column_index, value in enumerate(row, 1):
                if value != "":
                    cells[f"{get_column_letter(column_index)}{row_index}"] = CellData(value=value)
        return cells


class CsvReader:
    def read(self, path: Path, options: CompareOptions | None = None) -> CsvDocument:
        options = options or CompareOptions()
        try:
            with path.open("r", encoding=options.csv_encoding, newline="") as stream:
                return CsvDocument([row for row in csv.reader(stream, delimiter=options.csv_delimiter)])
        except LookupError as exc:
            raise WorkbookReadError(f"CSVの文字コード指定が不正です: {options.csv_encoding}") from exc
        except UnicodeDecodeError as exc:
            raise WorkbookReadError(f"CSVファイルを {options.csv_encoding} として読み取れません: {path}") from exc
        except OSError as exc:
            raise WorkbookReadError(f"CSVファイルを読み取れません: {path}") from exc
        except csv.Error as exc:
            raise WorkbookReadError(f"CSVファイルが読み取れません: {path}") from exc


class CsvComparer(Comparer[CsvDocument]):
    def __init__(self, table_comparer: ExcelComparer | None = None) -> None:
        self.table_comparer = table_comparer or ExcelComparer()

    def compare(
        self,
        old: CsvDocument,
        new: CsvDocument,
        options: CompareOptions,
        cancel_requested: CancelCheck | None = None,
    ) -> CompareResult:
        return self.table_comparer.compare(
            ExcelDocument({CSV_SHEET_NAME: old.cells}),
            ExcelDocument({CSV_SHEET_NAME: new.cells}),
            options,
            cancel_requested,
        )


class CsvReportWriter(ExcelReportWriter):
    def write(
        self,
        source_new: Path,
        output: Path,
        result: CompareResult,
        detailed: bool = True,
        cancel_requested: CancelCheck | None = None,
        options: CompareOptions | None = None,
    ) -> Path:
        workbook = None
        temporary_output: Path | None = None
        try:
            self._raise_if_cancelled(cancel_requested)
            output.parent.mkdir(parents=True, exist_ok=True)
            file_descriptor, temporary_name = tempfile.mkstemp(
                dir=output.parent,
                prefix=f".{output.stem}_",
                suffix=".tmp.xlsx",
            )
            os.close(file_descriptor)
            temporary_output = Path(temporary_name)

            csv_document = CsvReader().read(source_new, options)
            workbook = Workbook()
            csv_sheet = workbook.active
            csv_sheet.title = CSV_SHEET_NAME
            self._write_csv_sheet(csv_sheet, csv_document, cancel_requested)
            report = workbook.create_sheet(self._unique_report_name(workbook.sheetnames), 0)
            self._write_report(report, result, detailed, cancel_requested)
            self._highlight_differences(workbook, result, cancel_requested)
            self._raise_if_cancelled(cancel_requested)
            workbook.save(temporary_output)
            workbook.close()
            workbook = None
            os.replace(temporary_output, output)
            temporary_output = None
            result.output_path = output
            return output
        except (PermissionError, OSError, InvalidFileException, ValueError, csv.Error) as exc:
            raise OutputWriteError(f"出力ファイルを保存できません: {output}") from exc
        finally:
            if workbook is not None:
                workbook.close()
            if temporary_output is not None:
                with suppress(OSError):
                    temporary_output.unlink(missing_ok=True)

    @staticmethod
    def _write_csv_sheet(
        sheet: Worksheet,
        document: CsvDocument,
        cancel_requested: CancelCheck | None = None,
    ) -> None:
        for row_index, row in enumerate(document.rows, 1):
            raise_if_cancelled(cancel_requested)
            for column_index, value in enumerate(row, 1):
                sheet.cell(row_index, column_index, value)
