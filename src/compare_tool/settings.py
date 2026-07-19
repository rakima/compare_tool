from __future__ import annotations

import json
import os
from pathlib import Path

from .models import CompareAlgorithm, CompareOptions


class AppSettingsStore:
    """Persists small GUI preferences independently from comparison logic."""

    def __init__(self, path: Path | None = None) -> None:
        self.path = path or self._default_path()

    def load_last_save_dir(self) -> Path | None:
        try:
            data = self._read()
            value = data["last_save_dir"]
            if not isinstance(value, str):
                return None
            directory = Path(value)
            return directory if directory.is_dir() else None
        except (KeyError, TypeError, ValueError):
            return None

    def save_last_save_dir(self, directory: Path) -> None:
        data = self._read()
        data["last_save_dir"] = str(directory.resolve())
        self._write(data)

    def load_file_history(self) -> list[str]:
        history = self._read().get("file_history", [])
        if not isinstance(history, list):
            return []
        return [value for value in history if isinstance(value, str) and Path(value).is_file()]

    def save_file_history(self, paths: list[str], limit: int = 10) -> None:
        unique: list[str] = []
        for value in paths:
            if not value:
                continue
            normalized = str(Path(value).expanduser().resolve())
            if normalized not in unique:
                unique.append(normalized)
        data = self._read()
        data["file_history"] = unique[:limit]
        self._write(data)

    def clear_file_history(self) -> None:
        data = self._read()
        data["file_history"] = []
        self._write(data)

    def load_view_mode(self) -> str:
        value = self._read().get("view_mode")
        return value if value in {"detail", "summary"} else "detail"

    def save_view_mode(self, view_mode: str) -> None:
        if view_mode not in {"detail", "summary"}:
            return
        data = self._read()
        data["view_mode"] = view_mode
        self._write(data)

    def load_compare_options(self) -> CompareOptions:
        raw = self._read().get("compare_options", {})
        if not isinstance(raw, dict):
            return CompareOptions()
        defaults = CompareOptions()
        return CompareOptions(
            compare_values=self._bool(raw.get("compare_values"), defaults.compare_values),
            compare_formulas=self._bool(raw.get("compare_formulas"), defaults.compare_formulas),
            empty_string_equals_empty=self._bool(
                raw.get("empty_string_equals_empty"),
                defaults.empty_string_equals_empty,
            ),
            ignore_surrounding_whitespace=self._bool(
                raw.get("ignore_surrounding_whitespace"),
                defaults.ignore_surrounding_whitespace,
            ),
            ignore_case=self._bool(raw.get("ignore_case"), defaults.ignore_case),
            algorithm=self._algorithm(raw.get("algorithm"), defaults.algorithm),
            key_columns=self._key_columns(raw.get("key_columns")),
            csv_encoding=self._string(raw.get("csv_encoding"), defaults.csv_encoding),
            csv_delimiter=self._string(raw.get("csv_delimiter"), defaults.csv_delimiter),
            ignore_csv_blank_lines=self._bool(
                raw.get("ignore_csv_blank_lines"),
                defaults.ignore_csv_blank_lines,
            ),
            ignore_json_object_key_order=self._bool(
                raw.get("ignore_json_object_key_order"),
                defaults.ignore_json_object_key_order,
            ),
            ignore_json_array_order=self._bool(
                raw.get("ignore_json_array_order"),
                defaults.ignore_json_array_order,
            ),
            ignore_xml_attribute_order=self._bool(
                raw.get("ignore_xml_attribute_order"),
                defaults.ignore_xml_attribute_order,
            ),
            ignore_xml_blank_text=self._bool(
                raw.get("ignore_xml_blank_text"),
                defaults.ignore_xml_blank_text,
            ),
        )

    def save_compare_options(self, options: CompareOptions) -> None:
        data = self._read()
        data["compare_options"] = {
            "compare_values": options.compare_values,
            "compare_formulas": options.compare_formulas,
            "empty_string_equals_empty": options.empty_string_equals_empty,
            "ignore_surrounding_whitespace": options.ignore_surrounding_whitespace,
            "ignore_case": options.ignore_case,
            "algorithm": options.algorithm.value,
            "key_columns": list(options.key_columns),
            "csv_encoding": options.csv_encoding,
            "csv_delimiter": options.csv_delimiter,
            "ignore_csv_blank_lines": options.ignore_csv_blank_lines,
            "ignore_json_object_key_order": options.ignore_json_object_key_order,
            "ignore_json_array_order": options.ignore_json_array_order,
            "ignore_xml_attribute_order": options.ignore_xml_attribute_order,
            "ignore_xml_blank_text": options.ignore_xml_blank_text,
        }
        self._write(data)

    def _read(self) -> dict[str, object]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError, json.JSONDecodeError):
            return {}

    def _write(self, data: dict[str, object]) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError:
            # Saving a convenience preference must not turn a successful
            # workbook comparison into an application error.
            pass

    @staticmethod
    def _bool(value: object, default: bool) -> bool:
        return value if isinstance(value, bool) else default

    @staticmethod
    def _string(value: object, default: str) -> str:
        return value if isinstance(value, str) else default

    @staticmethod
    def _algorithm(value: object, default: CompareAlgorithm) -> CompareAlgorithm:
        if not isinstance(value, str):
            return default
        try:
            return CompareAlgorithm(value)
        except ValueError:
            return default

    @staticmethod
    def _key_columns(value: object) -> tuple[str, ...]:
        if not isinstance(value, list):
            return ()
        return tuple(item for item in value if isinstance(item, str))

    @staticmethod
    def _default_path() -> Path:
        base = os.environ.get("APPDATA")
        if base:
            return Path(base) / "compare_tool" / "settings.json"
        return Path.home() / ".config" / "compare_tool" / "settings.json"
