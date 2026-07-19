from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .comparer import CancelCheck, Comparer
from .errors import OperationCancelledError, WorkbookReadError
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
