from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Any


class DifferenceType(str, Enum):
    MODIFIED = "変更"
    ADDED = "追加"
    DELETED = "削除"
    ROW_ADDED = "行追加"
    ROW_DELETED = "行削除"
    SHEET_ADDED = "シート追加"
    SHEET_DELETED = "シート削除"


class CompareAlgorithm(str, Enum):
    CELL_COORDINATE = "cell_coordinate"
    ROW_LCS = "row_lcs"
    KEY_COLUMNS = "key_columns"


@dataclass(frozen=True, slots=True)
class CompareOptions:
    compare_values: bool = True
    compare_formulas: bool = True
    empty_string_equals_empty: bool = True
    ignore_surrounding_whitespace: bool = False
    ignore_case: bool = False
    algorithm: CompareAlgorithm = CompareAlgorithm.CELL_COORDINATE
    key_columns: tuple[str, ...] = ()
    csv_encoding: str = "utf-8-sig"
    csv_delimiter: str = ","


@dataclass(frozen=True, slots=True)
class Difference:
    kind: DifferenceType
    sheet: str
    cell: str | None = None
    old_value: Any = None
    new_value: Any = None
    value_changed: bool = False
    formula_changed: bool = False

    @property
    def can_link(self) -> bool:
        return self.kind in {DifferenceType.MODIFIED, DifferenceType.ADDED} and bool(self.cell)


@dataclass(slots=True)
class CompareResult:
    differences: list[Difference] = field(default_factory=list)
    output_path: Path | None = None

    def count(self, kind: DifferenceType) -> int:
        return sum(item.kind is kind for item in self.differences)

    @property
    def total(self) -> int:
        return len(self.differences)

    def summary(self) -> dict[DifferenceType, int]:
        return {kind: self.count(kind) for kind in DifferenceType}
