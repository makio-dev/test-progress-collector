# テスト進捗集計ツール v4

テストケースの予定・実績を集計し、進捗状況を可視化するExcelレポートを生成するツールです。

## 機能概要

- **ダッシュボード**: 本日のサマリー、チーム別進捗、進捗推移チャート、欠陥状況
- **欠陥ダッシュボード**: 欠陥の詳細分析（サマリー、対応状況別、緊急度別、業務機能分類、欠陥原因）
- **要対応一覧**: 遅延しているテストケースの一覧
- **進捗サマリー**: 日付×予定/実績の件数集計（チーム別シート）
- **欠陥サマリー**: 欠陥の検出・対応推移集計（全体＋チーム別シート）
- **欠陥詳細**: 欠陥一覧の全レコード詳細（全体＋チーム別シート）
- **明細シート**: 全テストケースの詳細一覧
- **祝日マスタ**: 営業日判定用の祝日管理

### 主な特徴

- ウィザード形式のGUI（tkinter）
- **複数フォルダの同時指定**（GUI・CLI 両対応／フォルダをまたいだ重複ファイルは自動除外）
- チーム名自動識別（ファイル名パターン: `-O-`:オンライン, `-B-`:バッチ, `-I-`:基盤, `-U-`:運用）
- サブフォルダを含む再帰的なファイル収集
- 対象ファイルは **ファイル名が `ITB-` で始まる** Excel のみ（無関係なファイルを自動除外）
- 基準日は**前営業日**（朝会で確定済みデータを表示するため）
- 欠陥一覧ファイルの取り込みと集計（チーム別）
- 欠陥詳細データ（テスト欠陥一覧シート）の読み取りと分析ダッシュボード
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
   - 「フォルダ追加」で**複数フォルダを登録**でき、登録済みフォルダは一覧表示・個別削除が可能です
   - 同じファイルが複数フォルダに含まれていても、絶対パスで重複排除され二重集計されません
2. **欠陥一覧ファイル選択**: チーム別の欠陥一覧ファイルを選択（任意）
3. **週範囲設定**: 週次集計の開始日・終了日を設定（デフォルト: 今週の月曜〜金曜）
4. **出力先設定**: 出力Excelファイルのパスを指定
5. **確認・実行**: 設定内容を確認して実行

### CLIモード

引数にフォルダパスを1つ以上渡すとCLIモードで動作します（引数なしで起動するとGUIウィザードになります）。

#### 複数フォルダ指定について

入力フォルダは**位置引数として何個でも並べて指定**できます（`-o` などのオプションより前後どちらに置いても構いません）。チームごとにフォルダが分かれている場合や、別ドライブ・別階層に散らばったテストケースをまとめて1つのレポートに集計したいときに便利です。

```bash
# 2フォルダを集計
python aggregate_test_results.py ./input_teamA ./input_teamB -o ./output/report.xlsx

# 3フォルダ以上もそのまま並べるだけ
python aggregate_test_results.py ./online ./batch ./infra ./ops -o ./output/report.xlsx

# 絶対パスと相対パスの混在もOK
python aggregate_test_results.py ./input /Volumes/share/テスト/運用 -o ./output/report.xlsx
```

複数フォルダ指定時の挙動：

- **再帰収集**: 各フォルダは既定でサブフォルダまで再帰的に走査されます（`--no-subfolders` で無効化）。
- **対象ファイルの絞り込み**: 各フォルダ内の `.xlsx` / `.xlsm` のうち、**ファイル名が `ITB-` で始まるもの**だけを集計対象とします（大文字小文字は区別しません）。Excel の一時ファイル（`~$` で始まる）や無関係なファイルは自動的に除外されます。さらに、シート名が `ITB-` で始まるシートを持たないファイルもスキップされます。
- **重複ファイルの自動排除**: フォルダの指定範囲が重なっていて同一ファイルが複数回ヒットしても、**絶対パスで重複判定**して一度しか集計しません（二重カウントされません）。
- **チーム振り分け**: フォルダではなく**ファイル名のパターン**（`-O-` / `-B-` / `-I-` / `-U-`）でチームを判定します。そのため、フォルダ構成に関係なくチーム別シートが正しく生成されます。
- **キャッシュ**: 差分更新キャッシュ（`.test_collector_cache.json`）は**出力ファイルと同じフォルダ**に作られ、全入力フォルダ横断で共有されます。前回から更新のないファイルは `[SKIP]` 表示でスキップされ高速化されます。

```bash
# 基本的な使い方（単一フォルダ）
python aggregate_test_results.py ./input -o ./output/report.xlsx

# サブフォルダを除外（指定フォルダ直下のみ集計）
python aggregate_test_results.py ./input -o ./output/report.xlsx --no-subfolders

# 週範囲を指定（スラッシュ形式）
python aggregate_test_results.py ./input -o ./output/report.xlsx --week-from 2026/03/01 --week-to 2026/03/07

# 週範囲を指定（スラッシュなし形式）
python aggregate_test_results.py ./input -o ./output/report.xlsx --week-from 20260301 --week-to 20260307

# 欠陥一覧ファイルを指定（必要なチームのみ指定可能）
python aggregate_test_results.py ./input -o ./output/report.xlsx \
    --defect-online ./input/defects/欠陥一覧_オンライン.xlsx \
    --defect-batch ./input/defects/欠陥一覧_バッチ.xlsx \
    --defect-infra ./input/defects/欠陥一覧_基盤.xlsx \
    --defect-ops ./input/defects/欠陥一覧_運用.xlsx

# 全オプション指定の例
python aggregate_test_results.py ./input -o ./output/report.xlsx \
    --week-from 2026/03/01 --week-to 2026/03/07 \
    --defect-online ./input/defects/欠陥一覧_オンライン.xlsx \
    --defect-batch ./input/defects/欠陥一覧_バッチ.xlsx
```

#### Windows (PowerShell / コマンドプロンプト)

```powershell
# 基本的な使い方（単一フォルダ）
python aggregate_test_results.py .\input -o .\output\report.xlsx

# 複数フォルダを指定
python aggregate_test_results.py .\input_teamA .\input_teamB -o .\output\report.xlsx

# サブフォルダを除外
python aggregate_test_results.py .\input -o .\output\report.xlsx --no-subfolders

# 週範囲を指定（スラッシュ形式）
python aggregate_test_results.py .\input -o .\output\report.xlsx --week-from 2026/03/01 --week-to 2026/03/07

# 週範囲を指定（スラッシュなし形式）
python aggregate_test_results.py .\input -o .\output\report.xlsx --week-from 20260301 --week-to 20260307

# 欠陥一覧ファイルを指定
python aggregate_test_results.py .\input -o .\output\report.xlsx `
    --defect-online .\input\defects\欠陥一覧_オンライン.xlsx `
    --defect-batch .\input\defects\欠陥一覧_バッチ.xlsx `
    --defect-infra .\input\defects\欠陥一覧_基盤.xlsx `
    --defect-ops .\input\defects\欠陥一覧_運用.xlsx

# EXEの場合も同様（相対パス・絶対パスどちらも可）
.\aggregate_test_results.exe .\input -o .\output\report.xlsx `
    --defect-online C:\data\defects\欠陥一覧_オンライン.xlsx `
    --defect-batch .\input\defects\欠陥一覧_バッチ.xlsx
```

### CLIオプション

| オプション | 説明 |
|-----------|------|
| `<input_folder> [input_folder2 ...]` | テストケースExcelファイルが格納されたフォルダ。**スペース区切りで複数指定可**（位置引数） |
| `-o, --output` | 出力ファイルパス（デフォルト: `./output/test_progress_{日時}.xlsx`） |
| `-s, --subfolders` | サブフォルダを含める（デフォルトで有効） |
| `--no-subfolders` | サブフォルダを含めない |
| `--week-from` | 週集計の開始日（YYYY/MM/DD または YYYYMMDD形式。デフォルト: 今週の月曜） |
| `--week-to` | 週集計の終了日（YYYY/MM/DD または YYYYMMDD形式。デフォルト: 今週の金曜） |
| `--defect-online` | 欠陥一覧ファイルのパス（オンラインチーム） |
| `--defect-batch` | 欠陥一覧ファイルのパス（バッチチーム） |
| `--defect-infra` | 欠陥一覧ファイルのパス（基盤チーム） |
| `--defect-ops` | 欠陥一覧ファイルのパス（運用チーム） |

> **パスの指定**: すべてのファイル・フォルダパスは**相対パス・絶対パスのどちらでも指定可能**です。
>
> **対象シートの接頭辞**は `ITB-` 固定です（CLIオプションでは変更できません。変更が必要な場合はソースの `SHEET_PREFIX` 定数を編集してください）。

## 入力ファイル形式

### テストケースファイル

#### 対象ファイル

- Excelファイル（`.xlsx`, `.xlsm`）
- **ファイル名が `ITB-` で始まる**もののみ対象（大文字小文字は区別しない。例: `ITB-O-001_ログイン認証.xlsx`）
- かつ、シート名が `ITB-` で始まるシートを持つこと（持たないファイルはスキップ）
- データ開始行は **19行目**（テスト実施者は `T7`、テスト検証者は `T8` セルから取得）

#### 必須列

| 列 | 内容 |
|----|------|
| C列 | テストID |
| O列 | 実施結果（初回） |
| P列 | 実施結果（再テスト） |
| Q列 | 実施者_予定日 |
| R列 | 実施者_実績日 |
| S列 | 検証者_予定日 |
| T列 | 検証者_実績日 |
| V列 | 欠陥内容／備考 |

> O列・P列・V列は明細シートの「実施結果（初回／再テスト）」「欠陥内容／備考」列として出力されます。

#### チーム名の自動識別

ファイル名に含まれるパターンでチーム名を自動判定：

| パターン | チーム名 |
|----------|----------|
| `-O-` | オンライン |
| `-B-` | バッチ |
| `-I-` | 基盤 |
| `-U-` | 運用 |
| その他 | その他 |

### 欠陥一覧ファイル

チーム別の欠陥一覧ファイルを指定することで、欠陥の検出・対応推移を集計できます。

#### ファイル要件

| 項目 | 仕様 |
|------|------|
| ファイル形式 | `.xlsx` |
| 必須シート名 | `欠陥発見・対応推移集計表` |
| ヘッダー行 | 10行目 |
| データ開始行 | 11行目 |

#### 必須列（欠陥発見・対応推移集計表）

| 列 | 内容 |
|----|------|
| B列 | No. |
| C列 | 日付 |
| D列 | 検出欠陥数 |
| E列 | 対応欠陥数 |
| F列 | 累積検出欠陥数 |
| G列 | 累積対応欠陥数 |
| H列 | 累積未対応欠陥数 |

#### テスト欠陥一覧シート（任意）

欠陥ダッシュボードを出力するには、欠陥一覧ファイルに `テスト欠陥一覧` シートが必要です。

| 項目 | 仕様 |
|------|------|
| シート名 | `テスト欠陥一覧` |
| ヘッダー行 | 8行目 |
| データ開始行 | 9行目 |
| 集計フラグ | AP列（1=欠陥として集計、0=非欠陥として除外） |

主な列: 欠陥ID(A)、対応状況(B)、件名(C)、発見日(D)、業務機能分類(G)、緊急度(M)、影響度(N)、調査予定日(O)、調査完了日(P)、欠陥原因(T)、対応予定日(AC)、対応日(AD)、横展開(AF-AI)、リリース(AK-AL)、検証日(AM)

#### 配置例

```
input/
└── defects/
    ├── 欠陥一覧_オンライン.xlsx
    ├── 欠陥一覧_バッチ.xlsx
    ├── 欠陥一覧_基盤.xlsx
    └── 欠陥一覧_運用.xlsx
```

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

# 欠陥一覧ファイルを指定（相対パス・絶対パスどちらも可）
.\aggregate_test_results.exe .\input -o .\output\report.xlsx `
    --defect-online .\input\defects\欠陥一覧_オンライン.xlsx `
    --defect-batch .\input\defects\欠陥一覧_バッチ.xlsx
```

### 注意事項

- `--windowed`オプションを付けても、CLIモードではコンソール出力が有効になります
- tkinterは標準ライブラリのため追加設定不要
- 初回起動時はWindows Defenderの警告が出る場合があります

## 出力ファイル構成

| シート名 | 内容 |
|----------|------|
| ダッシュボード | 本日の進捗サマリー、チャート、欠陥状況 |
| 欠陥ダッシュボード | 欠陥の詳細分析ダッシュボード（欠陥詳細データ指定時のみ） |
| 要対応一覧 | 遅延テストケース一覧 |
| 進捗サマリー_ALL | 全体の日次進捗 |
| 進捗サマリー_○○ | チーム別の日次進捗 |
| 欠陥サマリー_ALL | 全体の欠陥検出・対応推移（欠陥データ指定時のみ） |
| 欠陥サマリー_○○ | チーム別の欠陥検出・対応推移（欠陥データ指定時のみ） |
| 欠陥詳細_ALL | 全チームの欠陥詳細一覧（欠陥詳細データ指定時のみ） |
| 欠陥詳細_○○ | チーム別の欠陥詳細一覧（欠陥詳細データ指定時のみ） |
| 明細 | 全テストケースの詳細 |
| 祝日マスタ | 祝日一覧（編集可能） |

## PB曲線の生成（generate_pb_curve.py）

`generate_pb_curve.py` は、本ツールと同じ入力データから **PB曲線（信頼度成長曲線）** を生成する独立スクリプトです。1枚のグラフに次の2系統を重ねて、テストの進み具合と欠陥の出方を一目で評価できます。

- **P系（テスト消化）**: 未実施テストケースの残数バーンダウン（計画・実績）。`実施者_実績` をベースに集計。
- **B系（欠陥検出）**: 欠陥の累積検出数（実績・計画）、目標レンジ（ピンク帯）、基準日以降の予測。欠陥詳細の `発見日` をベースに集計。

### 特徴：生成後にExcel上で編集 → 自動再計算

生成されるExcelは **数式駆動** です。`入力データ`・`欠陥データ` シートは元データをコピーした編集可能リストで、グラフはこれらを直接参照しています。
特定の欠陥を集計から外したい場合は、**`欠陥データ` シートの該当行を削除して再オープンするだけ** で、累積欠陥カーブとグラフが自動的に再計算されます（係数や基準日も `パラメータ` シートで直接編集可能）。

### 使い方（CLI）

```bash
# 基本（係数・基準日は既定値＝前営業日／写真の係数）
python generate_pb_curve.py ./input -o ./output/pb_curve.xlsx

# 欠陥一覧ファイルを指定（B系を集計するために必要）
python generate_pb_curve.py ./input -o ./output/pb_curve.xlsx \
    --defect-online ./input/defects/欠陥一覧_オンライン.xlsx \
    --defect-batch  ./input/defects/欠陥一覧_バッチ.xlsx \
    --defect-infra  ./input/defects/欠陥一覧_基盤.xlsx \
    --defect-ops    ./input/defects/欠陥一覧_運用.xlsx

# 基準日・期間・B系係数・予測倍率を指定
python generate_pb_curve.py ./input -o ./output/pb_curve.xlsx \
    --pivot-date 2026-06-12 --start-date 2026-04-13 --end-date 2026-09-04 \
    --b-final-rate 0.0105 --b-lower-rate 0.0035 --b-upper-rate 0.0213 \
    --forecast-mult 0.0224
```

| オプション | 既定値 | 説明 |
|------------|--------|------|
| `-o, --output` | `./output/pb_curve.xlsx` | 出力ファイルパス |
| `--no-subfolders` | （再帰する） | サブフォルダを探索しない |
| `--pivot-date` | 前営業日 | 基準日。実績はこの日まで描画、以降は予測 |
| `--start-date` / `--end-date` | データの最小/最大日 | 分析対象期間 |
| `--b-final-rate` | `0.0105` | B系最終計画 係数（テストケース数×この値＝最終計画欠陥数） |
| `--b-lower-rate` / `--b-upper-rate` | `0.0035` / `0.0213` | B系目標帯の下限/上限 係数 |
| `--forecast-mult` | 実績から自動算出 | 基準日以降の欠陥発生見込み（欠陥/ケース） |
| `--total-case` | 収集件数 | テストケース総数 |
| `--defect-online/-batch/-infra/-ops` | なし | 欠陥一覧ファイル（チーム別） |

### 生成シート構成

| シート名 | 区分 | 内容 |
|----------|------|------|
| パラメータ | 入力（編集可） | 開始日/終了日/基準日/テストケース総数/B系係数/予測倍率。`B_Plan_Final` 等は係数からの数式 |
| 入力データ | 入力（編集可） | テストケースの実施予定日・実績日（P系の元データ） |
| 欠陥データ | 入力（編集可） | 欠陥の発見日（B系の元データ）。行削除で集計除外 |
| P_シリーズ | 参照計算 | 日次の予定/実績消化・未実施残数（COUNTIFS/SUM） |
| B_シリーズ | 参照計算 | 日次の検出/累積/目標帯/予測 |
| グラフ | 参照計算 | PB曲線（P系=左軸／B系=右軸、ピンク帯は目標レンジ） |

### EXE化（既存ツールと同名・別build）

ウィルス検査を通過させるため、**出力EXE名は既存ツールと同じ `aggregate_test_results.exe`** にします。本体EXEを上書きしないよう、`--name` と `--distpath` で **別フォルダに同名出力** します。

```powershell
pip install pyinstaller

# PB曲線ジェネレータを「aggregate_test_results.exe」という名前で dist_pb に出力
pyinstaller --onefile --windowed --name aggregate_test_results --distpath dist_pb generate_pb_curve.py
# → dist_pb\aggregate_test_results.exe（本体の dist\aggregate_test_results.exe とは別フォルダ・同名）
```

- 配布時は本体EXEと **別フォルダ** に置いてください（同一フォルダには同名で共存できません）。
- `generate_pb_curve.py` は `aggregate_test_results.py` の収集関数を import して再利用するため、ビルド時に同モジュールも自動的に取り込まれます。

## ライセンス

MIT License

## 作成者

テスト進捗集計ツール開発チーム
