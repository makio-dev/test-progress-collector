# テスト進捗集計ツール v4

テストケースの予定・実績を集計し、進捗状況を可視化するExcelレポートを生成するツールです。

## 機能概要

- **ダッシュボード**: 本日のサマリー、チーム別進捗、進捗推移チャート
- **要対応一覧**: 遅延しているテストケースの一覧
- **進捗サマリー**: 日付×予定/実績の件数集計（チーム別シート）
- **明細シート**: 全テストケースの詳細一覧
- **祝日マスタ**: 営業日判定用の祝日管理

### 主な特徴

- ウィザード形式のGUI（tkinter）
- チーム名自動識別（ファイル名パターン: `-O-`:オンライン, `-B-`:バッチ, `-I-`:基盤, `-U-`:運用）
- サブフォルダを含む再帰的なファイル収集
- 差分更新（キャッシュによる高速化）
- 条件付き書式による進捗の視覚化
- 累計・消化率の自動計算

## 必要環境

- Python 3.10以上
- 依存ライブラリ: openpyxl

## セットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/your-repo/test-progress-collector.git
cd test-progress-collector
```

### 2. Python仮想環境の作成

#### macOS / Linux

```bash
# 仮想環境の作成
python3 -m venv .venv

# 仮想環境の有効化
source .venv/bin/activate

# 依存ライブラリのインストール
pip install -r requirements.txt
```

#### Windows (PowerShell)

```powershell
# 仮想環境の作成
python -m venv .venv

# 仮想環境の有効化
.venv\Scripts\Activate.ps1

# 依存ライブラリのインストール
pip install -r requirements.txt
```

#### Windows (コマンドプロンプト)

```cmd
REM 仮想環境の作成
python -m venv .venv

REM 仮想環境の有効化
.venv\Scripts\activate.bat

REM 依存ライブラリのインストール
pip install -r requirements.txt
```

## 使い方

### GUIモード（推奨）

```bash
python aggregate_test_results.py
```

ウィザードが起動し、以下のステップで設定できます：
1. **入力フォルダ選択**: テストケースExcelファイルが格納されたフォルダを選択
2. **週範囲設定**: 週次集計の開始日・終了日を設定（デフォルト: 今日日付）
3. **出力先設定**: 出力Excelファイルのパスを指定
4. **確認・実行**: 設定内容を確認して実行

### CLIモード

#### macOS / Linux

```bash
# 基本的な使い方
python aggregate_test_results.py ./input -o ./output/report.xlsx

# サブフォルダを除外
python aggregate_test_results.py ./input -o ./output/report.xlsx --no-recursive

# 週範囲を指定（スラッシュ形式）
python aggregate_test_results.py ./input -o ./output/report.xlsx --week-from 2026/03/01 --week-to 2026/03/07

# 週範囲を指定（スラッシュなし形式）
python aggregate_test_results.py ./input -o ./output/report.xlsx --week-from 20260301 --week-to 20260307
```

#### Windows (PowerShell / コマンドプロンプト)

```powershell
# 基本的な使い方
python aggregate_test_results.py .\input -o .\output\report.xlsx

# サブフォルダを除外
python aggregate_test_results.py .\input -o .\output\report.xlsx --no-recursive

# 週範囲を指定（スラッシュ形式）
python aggregate_test_results.py .\input -o .\output\report.xlsx --week-from 2026/03/01 --week-to 2026/03/07

# 週範囲を指定（スラッシュなし形式）
python aggregate_test_results.py .\input -o .\output\report.xlsx --week-from 20260301 --week-to 20260307
```

### CLIオプション

| オプション | 説明 |
|-----------|------|
| `<input_folder>` | テストケースExcelファイルが格納されたフォルダ |
| `-o, --output` | 出力ファイルパス（デフォルト: `./output/テスト進捗集計_{日時}.xlsx`） |
| `--no-recursive` | サブフォルダを含めない |
| `--sheet-prefix` | 対象シートの接頭辞（デフォルト: `ITB`） |
| `--week-from` | 週集計の開始日（YYYY/MM/DD または YYYYMMDD形式） |
| `--week-to` | 週集計の終了日（YYYY/MM/DD または YYYYMMDD形式） |

## 入力ファイル形式

### 対象ファイル

- Excelファイル（`.xlsx`, `.xlsm`）
- シート名が`ITB`で始まるシートを対象

### 必須列

| 列 | 内容 |
|----|------|
| A列 | テストID |
| D列 | 実施者_予定日 |
| E列 | 実施者_実績日 |
| H列 | 検証者_予定日 |
| I列 | 検証者_実績日 |

### チーム名の自動識別

ファイル名に含まれるパターンでチーム名を自動判定：

| パターン | チーム名 |
|----------|----------|
| `-O-` | オンライン |
| `-B-` | バッチ |
| `-I-` | 基盤 |
| `-U-` | 運用 |
| その他 | その他 |

## EXE化（Windows向け配布）

PyInstallerを使用してスタンドアロンのEXEファイルを作成できます。

### 1. PyInstallerのインストール

```bash
pip install pyinstaller
```

### 2. EXEの作成

```powershell
# 推奨（GUI/CLI両対応）
pyinstaller --onefile --windowed aggregate_test_results.py

# アイコン付き
pyinstaller --onefile --windowed --icon=app.ico aggregate_test_results.py
```

### 3. 出力先

`dist\aggregate_test_results.exe` にEXEファイルが生成されます。

### 4. EXEの使い方

**1つのEXEでGUIモードとCLIモードの両方に対応しています。**

#### GUIモード（ダブルクリック）

EXEファイルをダブルクリックすると、ウィザード形式のGUIが起動します。
コンソールウィンドウは表示されません。

#### CLIモード（コマンドライン）

コマンドプロンプトやPowerShellから引数を付けて実行すると、CLIモードで動作します。
コンソールに進捗状況が出力されます。

```powershell
# 基本的な使い方
.\aggregate_test_results.exe .\input -o .\output\report.xlsx

# 週範囲を指定
.\aggregate_test_results.exe .\input -o .\output\report.xlsx --week-from 2026/03/01 --week-to 2026/03/07
```

### 注意事項

- `--windowed`オプションを付けても、CLIモードではコンソール出力が有効になります
- tkinterは標準ライブラリのため追加設定不要
- 初回起動時はWindows Defenderの警告が出る場合があります

## 出力ファイル構成

| シート名 | 内容 |
|----------|------|
| ダッシュボード | 本日の進捗サマリー、チャート |
| 要対応一覧 | 遅延テストケース一覧 |
| 進捗サマリー_ALL | 全体の日次進捗 |
| 進捗サマリー_○○ | チーム別の日次進捗 |
| 明細 | 全テストケースの詳細 |
| 祝日マスタ | 祝日一覧（編集可能） |

## ライセンス

MIT License

## 作成者

テスト進捗集計ツール開発チーム
