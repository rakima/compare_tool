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

大容量・疎なExcelの性能回帰テストは、通常テストと分けて次のコマンドで実行できます。

```powershell
python -m pytest -m performance -q
```

GitHub Actionsでは毎週日曜日3時（日本時間）に実行され、必要なときは手動でも開始できます。

## Windows向け配布

PyInstallerで単体の実行ファイルを作成できます。

```powershell
python -m pip install -e ".[dev]"
python -m PyInstaller compare_tool.spec --noconfirm --clean
```

作成された実行ファイルは `dist\compare_tool.exe` です。配布前には、別フォルダへコピーして起動し、サンプルExcelで比較できることを確認してください。

GitHub Actionsの `Build Windows App` は手動実行、または `v*` タグのpushでWindows実行ファイルをビルドし、`compare_tool-v<version>-windows` artifactとして保存します。

リリース前の確認項目は [docs/RELEASE_CHECKLIST.md](docs/RELEASE_CHECKLIST.md) にまとめています。

## 既知の制限

現在の対象は `.xlsx` のみです。`.xls`、`.xlsm`、CSV、JSON、XMLには対応していません。

比較対象はセル値、数式文字列、シート追加、シート削除です。書式、コメント、図形、画像、行高、列幅、テーブル定義は比較しません。

セル座標で比較するため、行追加・行削除があるExcelでは、以降のセルが大量の変更として検出される場合があります。LCSによる行追加・削除判定は将来対応予定です。

数式の再計算は行いません。数式の計算結果はExcelファイル内に保存済みの値を使用します。

## よくあるエラー

- `旧ファイルと新ファイルに同じファイルが指定されています。`: 異なる2つのExcelファイルを指定してください。
- `.xlsx ファイルを指定してください。`: 初期版では `.xlsx` のみ対応しています。
- `Excelファイルが破損しているか、読み取れません。`: Excelで開けるか確認し、必要なら別名保存してから再実行してください。
- `パスワード付きExcelは比較できません。`: パスワードを解除したコピーを指定してください。
- `出力ファイルを保存できません。`: 出力先ファイルをExcelで開いていないか、保存先フォルダへ書き込めるか確認してください。
