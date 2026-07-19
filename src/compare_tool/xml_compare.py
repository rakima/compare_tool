from __future__ import annotations

import os
import tempfile
from contextlib import suppress
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.worksheet import Worksheet

from .comparer import CancelCheck, Comparer
from .errors import OperationCancelledError, OutputWriteError, WorkbookReadError
from .excel import ExcelReportWriter
from .models import CompareOptions, CompareResult, Difference, DifferenceType

XML_SHEET_NAME = "XML"


@dataclass(frozen=True, slots=True)
class XmlDocument:
    root: ElementTree.Element


class XmlReader:
    def read(self, path: Path) -> XmlDocument:
        try:
            parser = ElementTree.XMLParser(target=ElementTree.TreeBuilder(insert_comments=False))
            with path.open("r", encoding="utf-8-sig") as stream:
                return XmlDocument(ElementTree.parse(stream, parser=parser).getroot())
        except UnicodeDecodeError as exc:
            raise WorkbookReadError(
                f"XMLファイルをUTF-8として読み取れません: {path}\n"
                "XMLはUTF-8 / UTF-8 BOM付きで保存してください。"
                "文字化けする場合は、エディタでUTF-8として保存し直してから再実行してください。"
            ) from exc
        except ElementTree.ParseError as exc:
            raise WorkbookReadError(
                f"XMLファイルの形式を読み取れません: {path}\n"
                f"{exc.position[0]}行 {exc.position[1]}列付近でXML構文エラーが発生しました: {exc}。\n"
                "開始タグと終了タグの対応、属性の引用符、特殊文字のエスケープを確認してください。"
            ) from exc
        except OSError as exc:
            raise WorkbookReadError(
                f"XMLファイルを読み取れません: {path}\n"
                "ファイルが存在するか、他のアプリで使用中ではないか、読み取り権限があるか確認してください。"
            ) from exc


class XmlComparer(Comparer[XmlDocument]):
    def compare(
        self,
        old: XmlDocument,
        new: XmlDocument,
        options: CompareOptions,
        cancel_requested: CancelCheck | None = None,
    ) -> CompareResult:
        differences: list[Difference] = []
        self._compare_element(
            f"/{self._display_tag(old.root.tag)}",
            old.root,
            new.root,
            options,
            differences,
            cancel_requested,
        )
        return CompareResult(differences)

    def _compare_element(
        self,
        path: str,
        old: ElementTree.Element,
        new: ElementTree.Element,
        options: CompareOptions,
        differences: list[Difference],
        cancel_requested: CancelCheck | None,
    ) -> None:
        self._raise_if_cancelled(cancel_requested)
        if old.tag != new.tag:
            differences.append(
                Difference(
                    DifferenceType.MODIFIED,
                    XML_SHEET_NAME,
                    path,
                    self._display_tag(old.tag),
                    self._display_tag(new.tag),
                    value_changed=True,
                )
            )
            return

        self._compare_attributes(path, old, new, options, differences, cancel_requested)
        self._compare_text(path, old, new, options, differences)
        self._compare_children(path, old, new, options, differences, cancel_requested)

    def _compare_attributes(
        self,
        path: str,
        old: ElementTree.Element,
        new: ElementTree.Element,
        options: CompareOptions,
        differences: list[Difference],
        cancel_requested: CancelCheck | None,
    ) -> None:
        if not options.ignore_xml_attribute_order and list(old.attrib) != list(new.attrib):
            differences.append(
                Difference(
                    DifferenceType.MODIFIED,
                    XML_SHEET_NAME,
                    f"{path}/@*",
                    list(old.attrib),
                    list(new.attrib),
                    value_changed=True,
                )
            )

        for name in sorted(old.attrib.keys() - new.attrib.keys()):
            self._raise_if_cancelled(cancel_requested)
            differences.append(Difference(DifferenceType.DELETED, XML_SHEET_NAME, f"{path}/@{name}", old.attrib[name]))
        for name in sorted(new.attrib.keys() - old.attrib.keys()):
            self._raise_if_cancelled(cancel_requested)
            differences.append(
                Difference(DifferenceType.ADDED, XML_SHEET_NAME, f"{path}/@{name}", None, new.attrib[name])
            )
        for name in sorted(old.attrib.keys() & new.attrib.keys()):
            self._raise_if_cancelled(cancel_requested)
            if not self._equal(old.attrib[name], new.attrib[name], options):
                differences.append(
                    Difference(
                        DifferenceType.MODIFIED,
                        XML_SHEET_NAME,
                        f"{path}/@{name}",
                        old.attrib[name],
                        new.attrib[name],
                        value_changed=True,
                    )
                )

    def _compare_text(
        self,
        path: str,
        old: ElementTree.Element,
        new: ElementTree.Element,
        options: CompareOptions,
        differences: list[Difference],
    ) -> None:
        old_text = self._text_value(old.text, options)
        new_text = self._text_value(new.text, options)
        if not self._equal(old_text, new_text, options):
            differences.append(
                Difference(
                    DifferenceType.MODIFIED,
                    XML_SHEET_NAME,
                    path,
                    old_text,
                    new_text,
                    value_changed=True,
                )
            )

    def _compare_children(
        self,
        path: str,
        old: ElementTree.Element,
        new: ElementTree.Element,
        options: CompareOptions,
        differences: list[Difference],
        cancel_requested: CancelCheck | None,
    ) -> None:
        old_children = list(old)
        new_children = list(new)
        old_child_paths = self._child_paths(path, old_children)
        new_child_paths = self._child_paths(path, new_children)

        common_length = min(len(old_children), len(new_children))
        for index in range(common_length):
            self._compare_element(
                old_child_paths[index],
                old_children[index],
                new_children[index],
                options,
                differences,
                cancel_requested,
            )
        for index in range(common_length, len(old_children)):
            self._raise_if_cancelled(cancel_requested)
            differences.append(
                Difference(
                    DifferenceType.DELETED,
                    XML_SHEET_NAME,
                    old_child_paths[index],
                    self._element_summary(old_children[index]),
                )
            )
        for index in range(common_length, len(new_children)):
            self._raise_if_cancelled(cancel_requested)
            differences.append(
                Difference(
                    DifferenceType.ADDED,
                    XML_SHEET_NAME,
                    new_child_paths[index],
                    None,
                    self._element_summary(new_children[index]),
                )
            )

    @classmethod
    def _child_paths(cls, parent_path: str, children: list[ElementTree.Element]) -> list[str]:
        counts: dict[str, int] = {}
        paths: list[str] = []
        for child in children:
            tag = cls._display_tag(child.tag)
            counts[tag] = counts.get(tag, 0) + 1
            paths.append(f"{parent_path}/{tag}[{counts[tag]}]")
        return paths

    @classmethod
    def _element_summary(cls, element: ElementTree.Element) -> str:
        return ElementTree.tostring(element, encoding="unicode", short_empty_elements=True)

    @staticmethod
    def _display_tag(tag: str) -> str:
        if tag.startswith("{") and "}" in tag:
            namespace, local_name = tag[1:].split("}", 1)
            return f"{local_name}{{{namespace}}}"
        return tag

    @classmethod
    def _text_value(cls, value: str | None, options: CompareOptions) -> str | None:
        if value is None:
            return None
        if options.ignore_xml_blank_text and value.strip() == "":
            return None
        return value

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


class XmlReportWriter(ExcelReportWriter):
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

            xml_document = XmlReader().read(source_new)
            workbook = Workbook()
            xml_sheet = workbook.active
            xml_sheet.title = XML_SHEET_NAME
            self._write_xml_sheet(xml_sheet, xml_document, cancel_requested)
            report = workbook.create_sheet(self._unique_report_name(workbook.sheetnames), 0)
            self._write_report(report, self._displayable_result(result), detailed, cancel_requested)
            self._remove_xml_links(report, result)
            self._write_xml_settings(report, options or CompareOptions())
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

    def _write_xml_sheet(
        self,
        sheet: Worksheet,
        document: XmlDocument,
        cancel_requested: CancelCheck | None = None,
    ) -> None:
        sheet["A1"] = "新XML"
        sheet["A1"].font = Font(bold=True, size=14)
        text = self._pretty_xml(document.root)
        for row_index, line in enumerate(text.splitlines(), 2):
            self._raise_if_cancelled(cancel_requested)
            self._write_cell(sheet, row_index, 1, line)
        sheet.column_dimensions["A"].width = 120

    @staticmethod
    def _pretty_xml(root: ElementTree.Element) -> str:
        copied = ElementTree.fromstring(ElementTree.tostring(root, encoding="unicode"))
        ElementTree.indent(copied, space="  ")
        return ElementTree.tostring(copied, encoding="unicode", short_empty_elements=True)

    @staticmethod
    def _displayable_result(result: CompareResult) -> CompareResult:
        return CompareResult(
            [
                Difference(
                    difference.kind,
                    difference.sheet,
                    difference.cell,
                    XmlReportWriter._display_xml_value(difference.old_value),
                    XmlReportWriter._display_xml_value(difference.new_value),
                    difference.value_changed,
                    difference.formula_changed,
                )
                for difference in result.differences
            ],
            result.output_path,
        )

    @staticmethod
    def _display_xml_value(value: object) -> object:
        if isinstance(value, list):
            return ", ".join(str(item) for item in value)
        return value

    @staticmethod
    def _remove_xml_links(sheet: Worksheet, result: CompareResult) -> None:
        for row_index in range(11, 11 + result.total):
            cell = sheet.cell(row_index, 6)
            cell.value = None
            cell.hyperlink = None
            cell.style = "Normal"

    @classmethod
    def _write_xml_settings(cls, sheet: Worksheet, options: CompareOptions) -> None:
        sheet["H1"] = "XML読み込み設定"
        sheet["H1"].font = Font(bold=True, size=14)
        rows = [
            ("文字コード", "UTF-8 / UTF-8 BOM"),
            ("比較位置", "XPath風パス"),
            ("属性順を無視", "はい" if options.ignore_xml_attribute_order else "いいえ"),
            ("空白のみテキストを無視", "はい" if options.ignore_xml_blank_text else "いいえ"),
        ]
        for row_index, (label, value) in enumerate(rows, 2):
            cls._write_cell(sheet, row_index, 8, label)
            cls._write_cell(sheet, row_index, 9, value)
        sheet.column_dimensions["H"].width = 22
        sheet.column_dimensions["I"].width = 24
