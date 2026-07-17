# compare_tool

2つの `.xlsx` ファイルをセル座標で比較し、新ファイルのコピーに比較結果シートと色付けを追加するGUIツールです。元ファイルは変更しません。

## セットアップと起動

Python 3.10以降で次を実行してください。

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e ".[test]"
compare-tool
```

または `python -m compare_tool` でも起動できます。

## 機能

- セル値、保存済み数式文字列、シート追加・削除の比較
- 空文字、空白、大文字小文字の正規化オプション
- 詳細一覧またはサマリーのみの結果シート
- 変更セルを黄色、追加セルを緑で表示
- 詳細一覧から変更・追加セルへの内部リンク
- ファイル選択およびドラッグ＆ドロップ

数式の再計算は行いません。「セル値を比較」はファイル内に保存済みの計算結果を使用します。Excelで未計算の数式は、数式比較をONにして確認してください。

## 設計

`Comparer` が形式非依存の比較戦略、`Difference` / `CompareResult` が共通結果モデルです。Excel固有処理は `ExcelReader`、`ExcelComparer`、`ExcelReportWriter` に分離しているため、将来は同じインターフェースでCSV・JSON・XML用の実装を追加できます。GUIは `CompareUseCase` のみを呼び出します。

## テスト

```powershell
python -m pip install -e ".[dev]"
python -m ruff check .
python -m ruff format --check .
python -m mypy
python -m pytest -q
```

同じ品質チェックはGitHub ActionsでもPython 3.10／3.12（Windows）に対して自動実行されます。

現在の対象は `.xlsx` のみです。書式、コメント、図形、画像、行高、列幅、テーブル定義は比較しません。
