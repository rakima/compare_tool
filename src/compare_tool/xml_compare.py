from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from .comparer import CancelCheck, Comparer
from .errors import OperationCancelledError, WorkbookReadError
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
