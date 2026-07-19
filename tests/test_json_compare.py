from __future__ import annotations

import json
from pathlib import Path

import pytest

from compare_tool.errors import OperationCancelledError, WorkbookReadError
from compare_tool.json_compare import JSON_SHEET_NAME, JsonComparer, JsonReader
from compare_tool.models import CompareOptions, DifferenceType


def write_json(path: Path, data: object) -> Path:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return path


def test_json_reader_reads_utf8_bom_file(tmp_path: Path) -> None:
    path = tmp_path / "input.json"
    path.write_text('\ufeff{"name": "compare_tool"}', encoding="utf-8")

    document = JsonReader().read(path)

    assert document.data == {"name": "compare_tool"}


def test_json_reader_reports_invalid_json(tmp_path: Path) -> None:
    path = tmp_path / "broken.json"
    path.write_text("{broken", encoding="utf-8")

    with pytest.raises(WorkbookReadError, match="JSONファイルの形式"):
        JsonReader().read(path)


def test_json_comparer_reports_modified_added_and_deleted_paths(tmp_path: Path) -> None:
    old = JsonReader().read(
        write_json(tmp_path / "old.json", {"name": "old", "removed": True, "items": [{"price": 100}]})
    )
    new = JsonReader().read(write_json(tmp_path / "new.json", {"name": "new", "added": 1, "items": [{"price": 120}]}))

    result = JsonComparer().compare(old, new, CompareOptions())

    differences = sorted((difference.cell, difference.kind) for difference in result.differences)
    assert differences == [
        ("$.added", DifferenceType.ADDED),
        ("$.items[0].price", DifferenceType.MODIFIED),
        ("$.name", DifferenceType.MODIFIED),
        ("$.removed", DifferenceType.DELETED),
    ]
    assert {difference.sheet for difference in result.differences} == {JSON_SHEET_NAME}


def test_json_comparer_reports_array_items_by_index(tmp_path: Path) -> None:
    old = JsonReader().read(write_json(tmp_path / "old.json", ["A", "B"]))
    new = JsonReader().read(write_json(tmp_path / "new.json", ["A", "X", "C"]))

    result = JsonComparer().compare(old, new, CompareOptions())

    differences = [
        (difference.kind, difference.cell, difference.old_value, difference.new_value)
        for difference in result.differences
    ]
    assert differences == [
        (DifferenceType.MODIFIED, "$[1]", "B", "X"),
        (DifferenceType.ADDED, "$[2]", None, "C"),
    ]


def test_json_comparer_uses_string_normalization_options(tmp_path: Path) -> None:
    old = JsonReader().read(write_json(tmp_path / "old.json", {"name": " Sample "}))
    new = JsonReader().read(write_json(tmp_path / "new.json", {"name": "sample"}))

    result = JsonComparer().compare(
        old,
        new,
        CompareOptions(ignore_surrounding_whitespace=True, ignore_case=True),
    )

    assert result.total == 0


def test_json_comparer_escapes_non_identifier_keys(tmp_path: Path) -> None:
    old = JsonReader().read(write_json(tmp_path / "old.json", {"display name": "old"}))
    new = JsonReader().read(write_json(tmp_path / "new.json", {"display name": "new"}))

    result = JsonComparer().compare(old, new, CompareOptions())

    assert result.differences[0].cell == "$['display name']"


def test_json_comparer_can_be_cancelled(tmp_path: Path) -> None:
    old = JsonReader().read(write_json(tmp_path / "old.json", {"name": "old"}))
    new = JsonReader().read(write_json(tmp_path / "new.json", {"name": "new"}))

    with pytest.raises(OperationCancelledError):
        JsonComparer().compare(old, new, CompareOptions(), cancel_requested=lambda: True)
