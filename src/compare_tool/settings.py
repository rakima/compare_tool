from __future__ import annotations

import json
import os
from pathlib import Path


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
    def _default_path() -> Path:
        base = os.environ.get("APPDATA")
        if base:
            return Path(base) / "compare_tool" / "settings.json"
        return Path.home() / ".config" / "compare_tool" / "settings.json"
