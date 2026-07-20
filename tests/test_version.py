from __future__ import annotations

import re
from pathlib import Path

from compare_tool import __version__


def test_package_version_matches_project_metadata() -> None:
    pyproject = Path("pyproject.toml").read_text(encoding="utf-8")
    match = re.search(r'^version = "([^"]+)"$', pyproject, re.MULTILINE)

    assert match is not None
    assert __version__ == match.group(1)
