from __future__ import annotations

from pathlib import Path

import tomllib

from compare_tool import __version__


def test_package_version_matches_project_metadata() -> None:
    pyproject = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))

    assert __version__ == pyproject["project"]["version"]
