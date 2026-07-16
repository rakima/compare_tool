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
            data = json.loads(self.path.read_text(encoding="utf-8"))
            directory = Path(data["last_save_dir"])
            return directory if directory.is_dir() else None
        except (OSError, KeyError, TypeError, ValueError, json.JSONDecodeError):
            return None

    def save_last_save_dir(self, directory: Path) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(
                json.dumps({"last_save_dir": str(directory.resolve())}, ensure_ascii=False, indent=2),
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
