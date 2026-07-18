"""End-to-end acceptance tests using real .xlsx files."""

from collections.abc import Mapping
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

from compare_tool.errors import InvalidInputError, OperationCancelledError, OutputWriteError, WorkbookReadError
from compare_tool.excel import ExcelReportWriter
from compare_tool.models import CompareOptions, CompareResult, Difference, DifferenceType
from compare_tool.usecase import CompareUseCase


def make_workbook(path: Path, sheets: Mapping[str, Mapping[str, object]]) -> Path:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, cells in sheets.items():
        sheet = workbook.create_sheet(name)
        for coordinate, value in cells.items():
            sheet[coordinate] = value
    workbook.save(path)
    workbook.close()
    return path


def compare(
    tmp_path: Path,
    old_sheets: Mapping[str, Mapping[str, object]],
    new_sheets: Mapping[str, Mapping[str, object]],
    *,
    detailed: bool = True,
    options: CompareOptions | None = None,
) -> tuple[CompareResult, Path]:
    old = make_workbook(tmp_path / "old.xlsx", old_sheets)
    new = make_workbook(tmp_path / "new.xlsx", new_sheets)
    output = tmp_path / "output.xlsx"
    result = CompareUseCase().execute(old, new, output, options or CompareOptions(), detailed)
    return result, output


def test_same_input_file_is_rejected(tmp_path: Path) -> None:
    source = make_workbook(tmp_path / "same.xlsx", {"Data": {"A1": "value"}})
    with pytest.raises(InvalidInputError, match="同じファイル"):
        CompareUseCase().execute(source, source, tmp_path / "output.xlsx", CompareOptions())


@pytest.mark.parametrize("bad_name", ["input.xls", "input.xlsm", "input.csv"])
def test_unsupported_input_extension_is_rejected(tmp_path: Path, bad_name: str) -> None:
    bad = tmp_path / bad_name
    bad.write_bytes(b"not used")
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {}})
    with pytest.raises(InvalidInputError, match=r"\.xlsx"):
        CompareUseCase().execute(bad, new, tmp_path / "output.xlsx", CompareOptions())


def test_missing_input_file_is_rejected(tmp_path: Path) -> None:
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {}})
    with pytest.raises(InvalidInputError, match="見つかりません"):
        CompareUseCase().execute(tmp_path / "missing.xlsx", new, tmp_path / "output.xlsx", CompareOptions())


def test_input_path_cannot_be_used_as_output(tmp_path: Path) -> None:
    old = make_workbook(tmp_path / "old.xlsx", {"Data": {"A1": "old"}})
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "new"}})
    with pytest.raises(InvalidInputError, match="異なるパス"):
        CompareUseCase().execute(old, new, new, CompareOptions())


def test_corrupted_xlsx_is_reported_as_read_error(tmp_path: Path) -> None:
    corrupt = tmp_path / "corrupt.xlsx"
    corrupt.write_bytes(b"this is not an xlsx archive")
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {}})
    with pytest.raises(WorkbookReadError, match="破損"):
        CompareUseCase().execute(corrupt, new, tmp_path / "output.xlsx", CompareOptions())


def test_unusable_output_parent_is_reported_as_write_error(tmp_path: Path) -> None:
    old = make_workbook(tmp_path / "old.xlsx", {"Data": {}})
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "new"}})
    parent_file = tmp_path / "not_a_directory"
    parent_file.write_text("file", encoding="utf-8")
    with pytest.raises(OutputWriteError, match="保存できません"):
        CompareUseCase().execute(old, new, parent_file / "output.xlsx", CompareOptions())


def test_existing_output_is_preserved_when_report_creation_fails(tmp_path: Path) -> None:
    class FailingWriter(ExcelReportWriter):
        def _write_report(
            self,
            sheet: object,
            result: CompareResult,
            detailed: bool,
            cancel_requested: object = None,
        ) -> None:
            raise ValueError("simulated report failure")

    source = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "new"}})
    output = tmp_path / "existing.xlsx"
    original_content = b"existing output must survive"
    output.write_bytes(original_content)

    with pytest.raises(OutputWriteError):
        FailingWriter().write(source, output, CompareResult())

    assert output.read_bytes() == original_content
    assert list(tmp_path.glob(".existing_*.tmp.xlsx")) == []


def test_successful_report_atomically_replaces_existing_output(tmp_path: Path) -> None:
    old = make_workbook(tmp_path / "old.xlsx", {"Data": {"A1": "old"}})
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "new"}})
    output = tmp_path / "existing.xlsx"
    output.write_bytes(b"previous output")

    CompareUseCase().execute(old, new, output, CompareOptions())

    workbook = load_workbook(output)
    assert workbook.sheetnames[0] == "比較結果"
    assert workbook["Data"]["A1"].value == "new"
    workbook.close()
    assert list(tmp_path.glob(".existing_*.tmp.xlsx")) == []


def test_cancelled_compare_does_not_create_output(tmp_path: Path) -> None:
    old = make_workbook(tmp_path / "old.xlsx", {"Data": {"A1": "old"}})
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "new"}})
    output = tmp_path / "output.xlsx"

    with pytest.raises(OperationCancelledError):
        CompareUseCase().execute(old, new, output, CompareOptions(), cancel_requested=lambda: True)

    assert not output.exists()


def test_cancelled_report_preserves_existing_output(tmp_path: Path) -> None:
    source = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "new"}})
    output = tmp_path / "existing.xlsx"
    original_content = b"existing output must survive cancellation"
    output.write_bytes(original_content)
    result = CompareResult(
        [
            Difference(DifferenceType.MODIFIED, "Data", "A1", "old", "new"),
            Difference(DifferenceType.ADDED, "Data", "B1", None, "added"),
        ]
    )

    with pytest.raises(OperationCancelledError):
        ExcelReportWriter().write(source, output, result, cancel_requested=lambda: True)

    assert output.read_bytes() == original_content
    assert list(tmp_path.glob(".existing_*.tmp.xlsx")) == []


def test_existing_report_sheet_names_are_not_overwritten(tmp_path: Path) -> None:
    sheets = {"Data": {"A1": "old"}, "比較結果": {"A1": "keep"}, "比較結果_1": {"A1": "keep too"}}
    new_sheets = {**sheets, "Data": {"A1": "new"}}
    _, output = compare(tmp_path, sheets, new_sheets)
    workbook = load_workbook(output)
    assert workbook.sheetnames[:3] == ["比較結果_2", "Data", "比較結果"]
    assert workbook["比較結果"]["A1"].value == "keep"
    assert workbook["比較結果_1"]["A1"].value == "keep too"
    workbook.close()


def test_link_escapes_apostrophe_in_sheet_name(tmp_path: Path) -> None:
    sheet_name = "社長's Data"
    _, output = compare(tmp_path, {sheet_name: {"A1": "old"}}, {sheet_name: {"A1": "new"}})
    workbook = load_workbook(output)
    assert workbook["比較結果"]["F9"].hyperlink.target == "#'社長''s Data'!A1"
    workbook.close()


def test_summary_mode_omits_detail_table_but_keeps_counts(tmp_path: Path) -> None:
    result, output = compare(tmp_path, {"Data": {"A1": "old"}}, {"Data": {"A1": "new", "B1": "added"}}, detailed=False)
    workbook = load_workbook(output)
    report = workbook["比較結果"]
    assert report["B2"].value == result.count(DifferenceType.MODIFIED) == 1
    assert report["B3"].value == result.count(DifferenceType.ADDED) == 1
    assert report["A8"].value is None
    workbook.close()


def test_formula_text_change_is_reported_without_recalculation(tmp_path: Path) -> None:
    result, output = compare(
        tmp_path,
        {"Data": {"A1": 1, "A2": 2, "B1": "=SUM(A1:A2)"}},
        {"Data": {"A1": 1, "A2": 2, "B1": "=SUM(A1:A1)"}},
    )
    difference = result.differences[0]
    assert difference.cell == "B1"
    assert difference.formula_changed is True
    workbook = load_workbook(output)
    assert workbook["比較結果"]["D9"].value == "'=SUM(A1:A2)"
    assert workbook["比較結果"]["E9"].value == "'=SUM(A1:A1)"
    workbook.close()


def test_formatting_only_cells_are_ignored(tmp_path: Path) -> None:
    old = make_workbook(tmp_path / "old.xlsx", {"Data": {"A1": "same"}})
    new = make_workbook(tmp_path / "new.xlsx", {"Data": {"A1": "same"}})
    workbook = load_workbook(new)
    workbook["Data"]["Z100"].fill = PatternFill("solid", fgColor="FF0000")
    workbook.save(new)
    workbook.close()
    result = CompareUseCase().execute(old, new, tmp_path / "output.xlsx", CompareOptions())
    assert result.total == 0


def test_sheet_addition_and_deletion_are_summarized(tmp_path: Path) -> None:
    result, output = compare(tmp_path, {"Common": {}, "Old only": {}}, {"Common": {}, "New only": {}})
    assert result.count(DifferenceType.SHEET_ADDED) == 1
    assert result.count(DifferenceType.SHEET_DELETED) == 1
    workbook = load_workbook(output)
    report = workbook["比較結果"]
    assert report["B5"].value == 1
    assert report["B6"].value == 1
    workbook.close()
