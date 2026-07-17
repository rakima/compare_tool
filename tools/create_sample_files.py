"""Create a pair of workbooks for manually checking compare_tool."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "samples"


def add_common_sheet(workbook: Workbook, *, new: bool) -> None:
    sheet = workbook.active
    sheet.title = "商品一覧"
    sheet.append(["商品ID", "商品名", "単価", "数量", "金額", "備考"])
    rows = (
        (
            ["P001", "ノートPC", 120000 if new else 100000, 2, "=C2*D2", "価格変更" if new else ""],
            ["P002", "マウス", 2500, 10 if new else 8, "=C3*D3", "数量変更"],
            ["P004", "Webカメラ", 8500, 4, "=C4*D4", "追加セル・行"],
        )
        if new
        else (
            ["P001", "ノートPC", 100000, 2, "=C2*D2", ""],
            ["P002", "マウス", 2500, 8, "=C3*D3", "数量変更"],
            ["P003", "キーボード", 6000, 5, "=C4*D4", "新ファイルでは削除"],
        )
    )
    for row in rows:
        sheet.append(row)

    # A formula-only change that keeps the displayed intent obvious.
    sheet["H1"] = "数式変更確認"
    sheet["H2"] = "=SUM(E2:E4)" if new else "=SUM(E2:E3)"
    if new:
        sheet["J10"] = "新ファイルで追加されたセル"
    else:
        sheet["J11"] = "新ファイルでは削除されたセル"

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    sheet.freeze_panes = "A2"
    for column, width in {"A": 12, "B": 18, "C": 12, "D": 10, "E": 14, "F": 22, "H": 18, "J": 28}.items():
        sheet.column_dimensions[column].width = width


def add_option_sheet(workbook: Workbook, *, new: bool) -> None:
    sheet = workbook.create_sheet("オプション確認")
    sheet.append(["確認項目", "値", "期待結果"])
    sheet.append(["前後スペース", " Sample " if not new else "Sample", "無視ONなら差分なし"])
    sheet.append(["大文字小文字", "EXCEL" if not new else "excel", "無視ONなら差分なし"])
    sheet.append(["空セル", "" if not new else None, "同一視ONなら差分なし"])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 24


def create_workbook(path: Path, *, new: bool) -> None:
    workbook = Workbook()
    add_common_sheet(workbook, new=new)
    add_option_sheet(workbook, new=new)
    unique_name = "新規シート" if new else "廃止シート"
    unique = workbook.create_sheet(unique_name)
    unique["A1"] = "新ファイルで追加されたシート" if new else "新ファイルでは削除されるシート"
    workbook.save(path)
    workbook.close()


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    create_workbook(OUTPUT_DIR / "比較サンプル_旧.xlsx", new=False)
    create_workbook(OUTPUT_DIR / "比較サンプル_新.xlsx", new=True)
    print(f"Created sample workbooks in {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
