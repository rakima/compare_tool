from __future__ import annotations

import json
import os
import tempfile
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

from .comparer import CancelCheck, Comparer
from .errors import OperationCancelledError, OutputWriteError, WorkbookReadError
from .excel import ExcelReportWriter
from .models import CompareOptions, CompareResult, Difference, DifferenceType

JSON_SHEET_NAME = "JSON"


@dataclass(frozen=True, slots=True)
class JsonDocument:
    data: Any


class JsonReader:
    def read(self, path: Path) -> JsonDocument:
        try:
            with path.open("r", encoding="utf-8-sig") as stream:
                return JsonDocument(json.load(stream))
        except UnicodeDecodeError as exc:
            raise WorkbookReadError(f"JSONファイルをUTF-8として読み取れません: {path}") from exc
        except json.JSONDecodeError as exc:
            raise WorkbookReadError(f"JSONファイルの形式を読み取れません: {path}") from exc
        except OSError as exc:
            raise WorkbookReadError(f"JSONファイルを読み取れません: {path}") from exc


class JsonComparer(Comparer[JsonDocument]):
    def compare(
        self,
        old: JsonDocument,
        new: JsonDocument,
        options: CompareOptions,
        cancel_requested: CancelCheck | None = None,
    ) -> CompareResult:
        differences: list[Difference] = []
        self._compare_value("$", old.data, new.data, options, differences, cancel_requested)
        return CompareResult(differences)

    def _compare_value(
        self,
        path: str,
        old: Any,
        new: Any,
        options: CompareOptions,
        differences: list[Difference],
        cancel_requested: CancelCheck | None,
    ) -> None:
        self._raise_if_cancelled(cancel_requested)
        if isinstance(old, dict) and isinstance(new, dict):
            self._compare_object(path, old, new, options, differences, cancel_requested)
            return
        if isinstance(old, list) and isinstance(new, list):
            self._compare_array(path, old, new, options, differences, cancel_requested)
            return
        if not self._equal(old, new, options):
            differences.append(
                Difference(
                    DifferenceType.MODIFIED,
                    JSON_SHEET_NAME,
                    path,
                    old,
                    new,
                    value_changed=True,
                )
            )

    def _compare_object(
        self,
        path: str,
        old: dict[str, Any],
        new: dict[str, Any],
        options: CompareOptions,
        differences: list[Difference],
        cancel_requested: CancelCheck | None,
    ) -> None:
        for key in sorted(old.keys() - new.keys()):
            self._raise_if_cancelled(cancel_requested)
            differences.append(
                Difference(
                    DifferenceType.DELETED,
                    JSON_SHEET_NAME,
                    self._object_path(path, key),
                    old[key],
                )
            )
        for key in sorted(new.keys() - old.keys()):
            self._raise_if_cancelled(cancel_requested)
            differences.append(
                Difference(DifferenceType.ADDED, JSON_SHEET_NAME, self._object_path(path, key), None, new[key])
            )
        for key in sorted(old.keys() & new.keys()):
            self._compare_value(
                self._object_path(path, key),
                old[key],
                new[key],
                options,
                differences,
                cancel_requested,
            )

    def _compare_array(
        self,
        path: str,
        old: list[Any],
        new: list[Any],
        options: CompareOptions,
        differences: list[Difference],
        cancel_requested: CancelCheck | None,
    ) -> None:
        common_length = min(len(old), len(new))
        for index in range(common_length):
            self._compare_value(f"{path}[{index}]", old[index], new[index], options, differences, cancel_requested)
        for index in range(common_length, len(old)):
            self._raise_if_cancelled(cancel_requested)
            differences.append(Difference(DifferenceType.DELETED, JSON_SHEET_NAME, f"{path}[{index}]", old[index]))
        for index in range(common_length, len(new)):
            self._raise_if_cancelled(cancel_requested)
            differences.append(Difference(DifferenceType.ADDED, JSON_SHEET_NAME, f"{path}[{index}]", None, new[index]))

    @staticmethod
    def _object_path(base: str, key: str) -> str:
        if key.isidentifier():
            return f"{base}.{key}"
        escaped = key.replace("\\", "\\\\").replace("'", "\\'")
        return f"{base}['{escaped}']"

    @staticmethod
    def _equal(left: Any, right: Any, options: CompareOptions) -> bool:
        def normalize(value: Any) -> Any:
            if options.empty_string_equals_empty and value == "":
                value = None
            if isinstance(value, str):
                if options.ignore_surrounding_whitespace:
                    value = value.strip()
                if options.ignore_case:
                    value = value.casefold()
            return value

        return normalize(left) == normalize(right)

    @staticmethod
    def _raise_if_cancelled(cancel_requested: CancelCheck | None) -> None:
        if cancel_requested is not None and cancel_requested():
            raise OperationCancelledError("比較をキャンセルしました。")


class JsonReportWriter(ExcelReportWriter):
    def write(
        self,
        source_new: Path,
        output: Path,
        result: CompareResult,
        detailed: bool = True,
        cancel_requested: CancelCheck | None = None,
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

            json_document = JsonReader().read(source_new)
            workbook = Workbook()
            json_sheet = workbook.active
            json_sheet.title = JSON_SHEET_NAME
            self._write_json_sheet(json_sheet, json_document, cancel_requested)
            report = workbook.create_sheet(self._unique_report_name(workbook.sheetnames), 0)
            self._write_report(report, self._displayable_result(result), detailed, cancel_requested)
            self._remove_json_links(report, result)
            self._write_json_settings(report)
            self._raise_if_cancelled(cancel_requested)
            workbook.save(temporary_output)
            workbook.close()
            workbook = None
            os.replace(temporary_output, output)
            temporary_output = None
            result.output_path = output
            return output
        except (PermissionError, OSError, InvalidFileException, ValueError, TypeError) as exc:
            raise OutputWriteError(f"出力ファイルを保存できません: {output}") from exc
        finally:
            if workbook is not None:
                workbook.close()
            if temporary_output is not None:
                with suppress(OSError):
                    temporary_output.unlink(missing_ok=True)

    def _write_json_sheet(
        self,
        sheet: Worksheet,
        document: JsonDocument,
        cancel_requested: CancelCheck | None = None,
    ) -> None:
        sheet["A1"] = "新JSON"
        sheet["A1"].font = Font(bold=True, size=14)
        text = json.dumps(document.data, ensure_ascii=False, indent=2)
        for row_index, line in enumerate(text.splitlines(), 2):
            self._raise_if_cancelled(cancel_requested)
            self._write_cell(sheet, row_index, 1, line)
        sheet.column_dimensions["A"].width = 100

    @staticmethod
    def _displayable_result(result: CompareResult) -> CompareResult:
        return CompareResult(
            [
                Difference(
                    difference.kind,
                    difference.sheet,
                    difference.cell,
                    JsonReportWriter._display_json_value(difference.old_value),
                    JsonReportWriter._display_json_value(difference.new_value),
                    difference.value_changed,
                    difference.formula_changed,
                )
                for difference in result.differences
            ],
            result.output_path,
        )

    @staticmethod
    def _display_json_value(value: object) -> object:
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False)
        return value

    @staticmethod
    def _remove_json_links(sheet: Worksheet, result: CompareResult) -> None:
        for row_index in range(11, 11 + result.total):
            cell = sheet.cell(row_index, 6)
            cell.value = None
            cell.hyperlink = None
            cell.style = "Normal"

    @classmethod
    def _write_json_settings(cls, sheet: Worksheet) -> None:
        sheet["H1"] = "JSON読み込み設定"
        sheet["H1"].font = Font(bold=True, size=14)
        rows = [
            ("文字コード", "UTF-8 / UTF-8 BOM"),
            ("比較位置", "JSON Path"),
            ("配列比較", "インデックス比較"),
        ]
        for row_index, (label, value) in enumerate(rows, 2):
            cls._write_cell(sheet, row_index, 8, label)
            cls._write_cell(sheet, row_index, 9, value)
        sheet.column_dimensions["H"].width = 18
        sheet.column_dimensions["I"].width = 24
