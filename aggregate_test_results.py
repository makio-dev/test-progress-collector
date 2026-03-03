"""
============================================================
  テスト予定・実績 集計スクリプト v4
============================================================
  概要：
    指定フォルダ内の全Excelファイルから
    「ITB」で始まるシートを対象に、
    実施者・検証者の予定/実績日付を読み取り、
    明細＋進捗サマリー＋祝日の3シート構成でExcelに出力します。

  出力構成：
    [進捗サマリー]   日付×予定/実績の件数集計（明細参照式+累計+進捗判定）
    [明細シート]     ファイル/シート/テストID単位の全レコード（テーブル形式）
    [祝日マスタ]     日本の祝日を管理（営業日判定用）

  主な機能：
    - ウィザード形式の使いやすいUI
    - チーム名自動識別（-O-:オンライン, -B-:バッチ, -I-:基盤, -U-:運用）
    - サブフォルダ含む再帰的なファイル収集
    - 差分更新（前回集計済みファイルはスキップ）
    - 日付別サマリーは明細シートを関数で参照
    - 進捗判定列（順調/遅延/完了/予定）
    - 条件付き書式による進捗の視覚化
    - 累計列の自動計算
    - 営業日/非営業日の識別
    - 基準日セルによる進捗基準の設定

  使い方：
    1. Python で実行: python aggregate_test_results.py
    ※ EXE化: pyinstaller --onefile --windowed aggregate_test_results.py

  EXE化時の注意：
    - --windowedオプションでコンソール非表示
    - tkinterは標準ライブラリなのでそのまま使用可能
============================================================
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
import os
import sys
from collections import defaultdict
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json

# ===================================================================
#  ===設定=== ここを環境に合わせて変更してください
# ===================================================================

# --- シートのフィルタ条件 ---
SHEET_PREFIX = "ITB"

# --- データ位置の設定 ---
COL_TEST_ID = 3  # C列
COL_JISSHI_YOTEI   = 17  # Q列: 実施者 予定
COL_JISSHI_JISSEKI  = 18  # R列: 実施者 実績
COL_KENSHO_YOTEI   = 19  # S列: 検証者 予定
COL_KENSHO_JISSEKI  = 20  # T列: 検証者 実績
DATA_START_ROW = 19

# --- チーム識別パターン ---
TEAM_PATTERNS = {
    "-O-": "オンライン",
    "-B-": "バッチ",
    "-I-": "基盤",
    "-U-": "運用",
}

# --- 日本の祝日（デフォルト: 2024-2026年） ---
DEFAULT_HOLIDAYS = [
    # 2024年
    "2024/01/01", "2024/01/08", "2024/02/11", "2024/02/12", "2024/02/23",
    "2024/03/20", "2024/04/29", "2024/05/03", "2024/05/04", "2024/05/05",
    "2024/05/06", "2024/07/15", "2024/08/11", "2024/08/12", "2024/09/16",
    "2024/09/22", "2024/09/23", "2024/10/14", "2024/11/03", "2024/11/04",
    "2024/11/23", "2024/12/23",
    # 2025年
    "2025/01/01", "2025/01/13", "2025/02/11", "2025/02/23", "2025/02/24",
    "2025/03/20", "2025/04/29", "2025/05/03", "2025/05/04", "2025/05/05",
    "2025/05/06", "2025/07/21", "2025/08/11", "2025/09/15", "2025/09/23",
    "2025/10/13", "2025/11/03", "2025/11/23", "2025/11/24",
    # 2026年
    "2026/01/01", "2026/01/12", "2026/02/11", "2026/02/23", "2026/03/20",
    "2026/04/29", "2026/05/03", "2026/05/04", "2026/05/05", "2026/05/06",
    "2026/07/20", "2026/08/11", "2026/09/21", "2026/09/22", "2026/09/23",
    "2026/10/12", "2026/11/03", "2026/11/23",
]

# ===================================================================
#  スタイル定義
# ===================================================================

TITLE_FONT = Font(name="游ゴシック", size=14, bold=True, color="1F4E79")
TITLE_FILL = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")

HEADER_FONT = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

DATA_FONT = Font(name="游ゴシック", size=10)
DATA_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
DATA_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
DATA_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

EVEN_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
CUMULATIVE_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

# 進捗状況の色
DONE_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")      # 緑：完了
NOT_DONE_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # 赤：未完了
DELAYED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # 赤：遅延
ONTRACK_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")   # 青：順調
HOLIDAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")   # グレー：非営業日
COMPLETED_ROW_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # グレー：完了行

# 基準日セルのスタイル
REF_DATE_FONT = Font(name="游ゴシック", size=12, bold=True, color="C00000")
REF_DATE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")


# ===================================================================
#  ウィザードUI
# ===================================================================

class WizardApp(tk.Tk):
    """ウィザード形式のメインアプリケーション"""

    def __init__(self):
        super().__init__()

        self.title("テスト進捗集計ツール v4")
        self.geometry("600x520")
        self.resizable(False, False)

        # 常に最前面に表示
        self.attributes("-topmost", True)

        # 結果を格納する変数
        self.result = None
        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.include_subfolders = tk.BooleanVar(value=True)
        self.update_mode = tk.StringVar(value="new")  # "new" or "update"

        # 現在のステップ
        self.current_step = 1

        # メインフレーム
        self.main_frame = ttk.Frame(self, padding=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # ヘッダー
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=(0, 20))

        self.title_label = ttk.Label(
            self.header_frame,
            text="テスト進捗集計ツール",
            font=("游ゴシック", 16, "bold")
        )
        self.title_label.pack()

        self.step_label = ttk.Label(
            self.header_frame,
            text="ステップ 1/3: 対象フォルダ選択",
            font=("游ゴシック", 11)
        )
        self.step_label.pack(pady=(5, 0))

        # コンテンツフレーム
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        # ボタンフレーム
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=(20, 0))

        self.back_btn = ttk.Button(self.button_frame, text="< 戻る", command=self.go_back)
        self.back_btn.pack(side=tk.LEFT)

        self.cancel_btn = ttk.Button(self.button_frame, text="キャンセル", command=self.cancel)
        self.cancel_btn.pack(side=tk.RIGHT, padx=(10, 0))

        self.next_btn = ttk.Button(self.button_frame, text="次へ >", command=self.go_next)
        self.next_btn.pack(side=tk.RIGHT)

        # 最初のステップを表示
        self.show_step(1)

        # ウィンドウを中央に配置
        self.center_window()

    def center_window(self):
        """ウィンドウを画面中央に配置"""
        self.update_idletasks()
        width = 600
        height = 520
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def clear_content(self):
        """コンテンツフレームをクリア"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def show_step(self, step):
        """指定されたステップを表示"""
        self.current_step = step
        self.clear_content()

        if step == 1:
            self.show_step1()
        elif step == 2:
            self.show_step2()
        elif step == 3:
            self.show_step3()

        # ボタンの状態を更新
        self.back_btn.config(state=tk.NORMAL if step > 1 else tk.DISABLED)
        self.next_btn.config(text="実行" if step == 3 else "次へ >")

    def show_step1(self):
        """ステップ1: 対象フォルダ選択"""
        self.step_label.config(text="ステップ 1/3: 対象フォルダ選択")

        # 説明
        desc = ttk.Label(
            self.content_frame,
            text="集計対象のExcelファイルが格納されているフォルダを選択してください。",
            wraplength=480
        )
        desc.pack(anchor=tk.W, pady=(0, 15))

        # フォルダ選択
        folder_frame = ttk.Frame(self.content_frame)
        folder_frame.pack(fill=tk.X, pady=10)

        ttk.Button(
            folder_frame,
            text="📁 フォルダを選択...",
            command=self.select_folder
        ).pack(side=tk.LEFT)

        self.folder_display = ttk.Label(
            folder_frame,
            text=self.folder_path.get() or "(未選択)",
            foreground="gray" if not self.folder_path.get() else "black"
        )
        self.folder_display.pack(side=tk.LEFT, padx=(10, 0))

        # サブフォルダオプション
        ttk.Checkbutton(
            self.content_frame,
            text="サブフォルダも含める（推奨）",
            variable=self.include_subfolders
        ).pack(anchor=tk.W, pady=(20, 0))

        # 説明追加
        note = ttk.Label(
            self.content_frame,
            text="※ サブフォルダを含めると、選択したフォルダ配下の全てのExcelファイルが対象になります。",
            foreground="gray",
            wraplength=480
        )
        note.pack(anchor=tk.W, pady=(5, 0))

    def show_step2(self):
        """ステップ2: 出力設定"""
        self.step_label.config(text="ステップ 2/3: 出力設定")

        # 説明
        desc = ttk.Label(
            self.content_frame,
            text="集計結果の出力方法を選択してください。",
            wraplength=480
        )
        desc.pack(anchor=tk.W, pady=(0, 15))

        # モード選択
        mode_frame = ttk.LabelFrame(self.content_frame, text="出力モード", padding=10)
        mode_frame.pack(fill=tk.X, pady=10)

        ttk.Radiobutton(
            mode_frame,
            text="新規作成",
            variable=self.update_mode,
            value="new"
        ).pack(anchor=tk.W)

        ttk.Label(
            mode_frame,
            text="   新しいExcelファイルを作成します",
            foreground="gray"
        ).pack(anchor=tk.W)

        ttk.Radiobutton(
            mode_frame,
            text="既存ファイルを更新（上書き）",
            variable=self.update_mode,
            value="update"
        ).pack(anchor=tk.W, pady=(10, 0))

        ttk.Label(
            mode_frame,
            text="   既存のExcelファイルを選択し、最新データで上書きします",
            foreground="gray"
        ).pack(anchor=tk.W)

        # ファイル選択
        file_frame = ttk.Frame(self.content_frame)
        file_frame.pack(fill=tk.X, pady=(20, 10))

        ttk.Button(
            file_frame,
            text="💾 保存先を選択...",
            command=self.select_output
        ).pack(side=tk.LEFT)

        self.output_display = ttk.Label(
            file_frame,
            text=self.output_path.get() or "(未選択)",
            foreground="gray" if not self.output_path.get() else "black"
        )
        self.output_display.pack(side=tk.LEFT, padx=(10, 0))

    def show_step3(self):
        """ステップ3: 確認"""
        self.step_label.config(text="ステップ 3/3: 確認")

        # 説明
        desc = ttk.Label(
            self.content_frame,
            text="以下の設定で集計を実行します。内容を確認してください。",
            wraplength=480
        )
        desc.pack(anchor=tk.W, pady=(0, 15))

        # 設定内容表示
        confirm_frame = ttk.LabelFrame(self.content_frame, text="設定内容", padding=15)
        confirm_frame.pack(fill=tk.X, pady=10)

        # フォルダ
        row1 = ttk.Frame(confirm_frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="対象フォルダ:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Label(row1, text=self.folder_path.get(), wraplength=350).pack(side=tk.LEFT)

        # サブフォルダ
        row2 = ttk.Frame(confirm_frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="サブフォルダ:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Label(row2, text="含める" if self.include_subfolders.get() else "含めない").pack(side=tk.LEFT)

        # 出力先
        row3 = ttk.Frame(confirm_frame)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="出力先:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Label(row3, text=self.output_path.get(), wraplength=350).pack(side=tk.LEFT)

        # モード
        row4 = ttk.Frame(confirm_frame)
        row4.pack(fill=tk.X, pady=3)
        ttk.Label(row4, text="モード:", width=15, anchor=tk.W).pack(side=tk.LEFT)
        mode_text = "新規作成" if self.update_mode.get() == "new" else "既存ファイル更新"
        ttk.Label(row4, text=mode_text).pack(side=tk.LEFT)

        # 注意書き
        note = ttk.Label(
            self.content_frame,
            text="※ 「実行」をクリックすると集計処理を開始します。\n※ 前回集計済みのファイルは自動的にスキップされます。",
            foreground="gray",
            wraplength=480
        )
        note.pack(anchor=tk.W, pady=(15, 0))

    def select_folder(self):
        """フォルダ選択ダイアログ"""
        folder = filedialog.askdirectory(title="対象フォルダを選択")
        if folder:
            self.folder_path.set(folder)
            self.folder_display.config(text=folder, foreground="black")

    def select_output(self):
        """出力ファイル選択ダイアログ"""
        if self.update_mode.get() == "new":
            path = filedialog.asksaveasfilename(
                title="保存先を選択",
                defaultextension=".xlsx",
                filetypes=[("Excel ファイル", "*.xlsx")],
                initialfile="テスト進捗集計.xlsx"
            )
        else:
            path = filedialog.askopenfilename(
                title="更新するファイルを選択",
                filetypes=[("Excel ファイル", "*.xlsx")]
            )

        if path:
            self.output_path.set(path)
            self.output_display.config(text=path, foreground="black")

    def go_back(self):
        """前のステップへ"""
        if self.current_step > 1:
            self.show_step(self.current_step - 1)

    def go_next(self):
        """次のステップへ / 実行"""
        if self.current_step == 1:
            if not self.folder_path.get():
                messagebox.showwarning("入力エラー", "対象フォルダを選択してください。")
                return
            self.show_step(2)

        elif self.current_step == 2:
            if not self.output_path.get():
                messagebox.showwarning("入力エラー", "出力先を選択してください。")
                return
            self.show_step(3)

        elif self.current_step == 3:
            self.execute()

    def execute(self):
        """集計を実行"""
        self.result = {
            "folder_path": self.folder_path.get(),
            "output_path": self.output_path.get(),
            "include_subfolders": self.include_subfolders.get(),
            "update_mode": self.update_mode.get(),
        }
        self.destroy()

    def cancel(self):
        """キャンセル"""
        self.result = None
        self.destroy()


def run_wizard():
    """ウィザードを実行して設定を取得"""
    app = WizardApp()
    app.mainloop()
    return app.result


# ===================================================================
#  ユーティリティ関数
# ===================================================================

def identify_team(filename):
    """ファイル名からチーム名を識別"""
    for pattern, team_name in TEAM_PATTERNS.items():
        if pattern in filename.upper():
            return team_name
    return "その他"


def _to_date(val):
    """セル値を日付文字列に変換（日付でなければNone）"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y/%m/%d")
    return None


def _to_date_obj(val):
    """セル値をdatetimeオブジェクトに変換"""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        try:
            return datetime.strptime(val, "%Y/%m/%d")
        except ValueError:
            return None
    return None


def generate_date_range(start_date, end_date):
    """開始日から終了日までの日付リストを生成"""
    dates = []
    current = start_date
    while current <= end_date:
        dates.append(current)
        current += timedelta(days=1)
    return dates


def is_weekend(date_obj):
    """土日判定"""
    return date_obj.weekday() >= 5  # 5=土曜, 6=日曜


def load_cache(cache_file):
    """キャッシュファイルを読み込む"""
    if cache_file and os.path.exists(cache_file):
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_cache(cache_file, cache_data):
    """キャッシュファイルを保存する"""
    if cache_file:
        try:
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


# ===================================================================
#  データ収集
# ===================================================================

def collect_data(folder_path, cache_file=None, include_subfolders=True):
    """フォルダ内の全Excelファイル → ITBシート → 全テストケースを収集"""

    records = []
    file_count = 0
    sheet_count = 0
    skipped_count = 0

    # キャッシュ読み込み
    cached_data = load_cache(cache_file)
    new_cache = {}

    # Excelファイルを収集
    excel_files = []
    if include_subfolders:
        for root, dirs, files in os.walk(folder_path):
            for filename in sorted(files):
                if filename.endswith(".xlsx") and not filename.startswith("~$"):
                    excel_files.append(os.path.join(root, filename))
    else:
        for filename in sorted(os.listdir(folder_path)):
            if filename.endswith(".xlsx") and not filename.startswith("~$"):
                excel_files.append(os.path.join(folder_path, filename))

    for filepath in excel_files:
        filename = os.path.basename(filepath)
        relative_path = os.path.relpath(filepath, folder_path)
        file_mtime = os.path.getmtime(filepath)

        # 差分チェック: ファイルの更新日時が変わっていなければキャッシュを使用
        if filepath in cached_data:
            cached_entry = cached_data[filepath]
            # 新形式: {mtime: ..., records: [...]} / 旧形式: float (mtime直接)
            if isinstance(cached_entry, dict) and cached_entry.get('mtime') == file_mtime:
                # キャッシュからレコードを復元
                cached_records = cached_entry.get('records', [])
                records.extend(cached_records)
                new_cache[filepath] = cached_entry
                skipped_count += 1
                print(f"  ⏭ {relative_path} (キャッシュ使用: {len(cached_records)}件)")
                continue
            # 旧形式のキャッシュは無視して再処理

        print(f"  📄 {relative_path}")

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True)
            file_count += 1

            # ファイル名からチーム名を識別
            team_name = identify_team(filename)
            file_records = []

            for ws_name in wb.sheetnames:
                if not ws_name.upper().startswith(SHEET_PREFIX.upper()):
                    continue

                ws = wb[ws_name]
                sheet_count += 1
                case_count = 0

                for row in range(DATA_START_ROW, ws.max_row + 1):
                    test_id = ws.cell(row=row, column=COL_TEST_ID).value
                    if not test_id:
                        continue

                    jisshi_yotei   = ws.cell(row=row, column=COL_JISSHI_YOTEI).value
                    jisshi_jisseki = ws.cell(row=row, column=COL_JISSHI_JISSEKI).value
                    kensho_yotei   = ws.cell(row=row, column=COL_KENSHO_YOTEI).value
                    kensho_jisseki = ws.cell(row=row, column=COL_KENSHO_JISSEKI).value

                    record = {
                        "ファイル名": relative_path,  # 相対パスで記録
                        "ファイルパス": filepath,
                        "シート名": ws_name,
                        "チーム名": team_name,
                        "テストID": str(test_id),
                        "実施者_予定": _to_date(jisshi_yotei),
                        "実施者_実績": _to_date(jisshi_jisseki),
                        "検証者_予定": _to_date(kensho_yotei),
                        "検証者_実績": _to_date(kensho_jisseki),
                    }
                    file_records.append(record)
                    case_count += 1

                print(f"     ✅ {ws_name} ({case_count}件) [チーム: {team_name}]")

            wb.close()

            # キャッシュに保存
            new_cache[filepath] = {
                'mtime': file_mtime,
                'records': file_records
            }
            records.extend(file_records)

        except Exception as e:
            print(f"     ⚠ エラー: {e}")

    # キャッシュを保存
    save_cache(cache_file, new_cache)

    print(f"\n  処理完了: {file_count}ファイル処理, {skipped_count}ファイルスキップ, {sheet_count}シート, {len(records)}レコード")
    return records


# ===================================================================
#  Excel出力
# ===================================================================

def write_excel(records, output_path, holidays=None):
    """明細シート＋進捗サマリー（チーム別）＋祝日マスタをExcelに出力"""

    if holidays is None:
        holidays = DEFAULT_HOLIDAYS

    wb = openpyxl.Workbook()

    # --- 祝日マスタシート ---
    ws_holiday = wb.active
    ws_holiday.title = "祝日マスタ"
    _write_holiday_sheet(ws_holiday, holidays)

    # --- 明細シート ---
    ws_detail = wb.create_sheet("明細")
    detail_data_start_row = _write_detail_sheet(ws_detail, records)

    # --- チーム別にレコードを分類 ---
    team_records = defaultdict(list)
    for rec in records:
        team_records[rec["チーム名"]].append(rec)

    # チームリスト（ALL + 実際のチーム）
    teams_in_data = sorted(team_records.keys())

    # --- 進捗サマリーシート（ALL） ---
    ws_summary_all = wb.create_sheet("進捗サマリー_ALL")
    _write_summary_sheet(ws_summary_all, records, detail_data_start_row, len(records), holidays, "ALL")

    # --- 進捗サマリーシート（チーム別） ---
    for team_name in teams_in_data:
        team_recs = team_records[team_name]
        sheet_name = f"進捗サマリー_{team_name}"
        ws_team = wb.create_sheet(sheet_name)
        _write_summary_sheet(ws_team, team_recs, detail_data_start_row, len(records), holidays, team_name)

    # シートの順序を調整（進捗サマリー_ALLを先頭に）
    wb.move_sheet(ws_summary_all, offset=-(len(teams_in_data) + 2))

    # --- 保存 ---
    wb.save(output_path)
    print(f"\n  ✅ 出力完了: {output_path}")
    print(f"     明細シート: {len(records)}件")
    print(f"     サマリーシート: ALL + {len(teams_in_data)}チーム")


def _write_holiday_sheet(ws, holidays):
    """祝日マスタシートを作成"""

    # タイトル
    ws.merge_cells('A1:C1')
    title_cell = ws['A1']
    title_cell.value = "祝日マスタ"
    title_cell.font = TITLE_FONT
    title_cell.fill = TITLE_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25

    # 説明
    ws['A2'] = "※ 祝日を追加・編集してください。日付形式: YYYY/MM/DD"
    ws['A2'].font = Font(name="游ゴシック", size=9, color="666666")

    # ヘッダー
    headers = ["日付", "祝日名（任意）", "備考"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    # 祝日データ
    for i, holiday in enumerate(holidays):
        row = i + 5
        ws.cell(row=row, column=1, value=holiday).border = THIN_BORDER
        ws.cell(row=row, column=2, value="").border = THIN_BORDER
        ws.cell(row=row, column=3, value="").border = THIN_BORDER

    # 列幅
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25


def _write_detail_sheet(ws, records):
    """明細シートを作成（テーブル形式）"""

    # タイトル (A1)
    ws.merge_cells('A1:L1')
    title_cell = ws['A1']
    title_cell.value = "テスト進捗明細"
    title_cell.font = TITLE_FONT
    title_cell.fill = TITLE_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25

    # 基準日ラベルと値 (K2, L2)
    ws['K2'] = "基準日:"
    ws['K2'].font = Font(name="游ゴシック", size=11, bold=True)
    ws['K2'].alignment = DATA_ALIGN_RIGHT

    ws['L2'] = datetime.now().strftime("%Y/%m/%d")
    ws['L2'].font = REF_DATE_FONT
    ws['L2'].fill = REF_DATE_FILL
    ws['L2'].alignment = DATA_ALIGN_CENTER
    ws['L2'].border = THIN_BORDER

    # ヘッダー行 (row 4)
    header_row = 4
    detail_headers = [
        "No.", "ファイル名", "シート名", "チーム名", "テストID",
        "実施者_予定", "実施者_実績", "実施者_状況",
        "検証者_予定", "検証者_実績", "検証者_状況",
        "進捗状況",
    ]

    for col, header in enumerate(detail_headers, 1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    # データ行
    data_start_row = header_row + 1
    for i, rec in enumerate(records):
        row = data_start_row + i

        # 実施者状況（実績があれば完了、なければ予定日と基準日を比較）
        jisshi_status = '=IF(G{row}<>"","完了",IF(F{row}="","－",IF(F{row}<=$L$2,"遅延","予定")))'.format(row=row)
        # 検証者状況
        kensho_status = '=IF(J{row}<>"","完了",IF(I{row}="","－",IF(I{row}<=$L$2,"遅延","予定")))'.format(row=row)
        # 全体進捗
        overall_status = '=IF(AND(H{row}="完了",K{row}="完了"),"完了",IF(OR(H{row}="遅延",K{row}="遅延"),"遅延","進行中"))'.format(row=row)

        values = [
            i + 1,
            rec["ファイル名"],
            rec["シート名"],
            rec["チーム名"],
            rec["テストID"],
            rec["実施者_予定"] if rec["実施者_予定"] else "",
            rec["実施者_実績"] if rec["実施者_実績"] else "",
            jisshi_status,
            rec["検証者_予定"] if rec["検証者_予定"] else "",
            rec["検証者_実績"] if rec["検証者_実績"] else "",
            kensho_status,
            overall_status,
        ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER

            if col == 1:
                cell.alignment = DATA_ALIGN_CENTER
            elif col in (4, 5, 6, 7, 8, 9, 10, 11, 12):
                cell.alignment = DATA_ALIGN_CENTER
            else:
                cell.alignment = DATA_ALIGN_LEFT

    # テーブル作成
    if records:
        table_ref = f"A{header_row}:L{data_start_row + len(records) - 1}"
        table = Table(displayName="明細テーブル", ref=table_ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    # 条件付き書式: 状況列（H, K, L列）
    last_row = data_start_row + len(records) - 1 if records else data_start_row

    # 完了 = 緑
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(color="006100")
    ws.conditional_formatting.add(
        f"H{data_start_row}:H{last_row}",
        FormulaRule(formula=[f'H{data_start_row}="完了"'], fill=green_fill, font=green_font)
    )
    ws.conditional_formatting.add(
        f"K{data_start_row}:K{last_row}",
        FormulaRule(formula=[f'K{data_start_row}="完了"'], fill=green_fill, font=green_font)
    )
    ws.conditional_formatting.add(
        f"L{data_start_row}:L{last_row}",
        FormulaRule(formula=[f'L{data_start_row}="完了"'], fill=green_fill, font=green_font)
    )

    # 遅延 = 赤
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006")
    ws.conditional_formatting.add(
        f"H{data_start_row}:H{last_row}",
        FormulaRule(formula=[f'H{data_start_row}="遅延"'], fill=red_fill, font=red_font)
    )
    ws.conditional_formatting.add(
        f"K{data_start_row}:K{last_row}",
        FormulaRule(formula=[f'K{data_start_row}="遅延"'], fill=red_fill, font=red_font)
    )
    ws.conditional_formatting.add(
        f"L{data_start_row}:L{last_row}",
        FormulaRule(formula=[f'L{data_start_row}="遅延"'], fill=red_fill, font=red_font)
    )

    # 予定 = 青
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    blue_font = Font(color="1F4E79")
    ws.conditional_formatting.add(
        f"H{data_start_row}:H{last_row}",
        FormulaRule(formula=[f'H{data_start_row}="予定"'], fill=blue_fill, font=blue_font)
    )
    ws.conditional_formatting.add(
        f"K{data_start_row}:K{last_row}",
        FormulaRule(formula=[f'K{data_start_row}="予定"'], fill=blue_fill, font=blue_font)
    )

    # 進行中 = オレンジ
    orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    orange_font = Font(color="C65911")
    ws.conditional_formatting.add(
        f"L{data_start_row}:L{last_row}",
        FormulaRule(formula=[f'L{data_start_row}="進行中"'], fill=orange_fill, font=orange_font)
    )

    # 列幅設定
    detail_widths = [6, 28, 18, 12, 16, 12, 12, 10, 12, 12, 10, 10]
    for i, w in enumerate(detail_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = f"A{data_start_row}"

    return data_start_row


def _write_summary_sheet(ws, records, detail_start_row, total_record_count, holidays, team_name="ALL"):
    """進捗サマリーシートを作成（明細参照式＋累計＋進捗判定:実施/検証別）"""

    # 日付範囲を取得（対象レコードから）
    all_dates = []
    for rec in records:
        for key in ["実施者_予定", "実施者_実績", "検証者_予定", "検証者_実績"]:
            if rec[key]:
                date_obj = _to_date_obj(rec[key])
                if date_obj:
                    all_dates.append(date_obj)

    if not all_dates:
        ws['A1'] = "データがありません"
        return

    min_date = min(all_dates)
    max_date = max(all_dates)
    date_range = generate_date_range(min_date, max_date)

    # 基準日（今日）
    today_str = datetime.now().strftime("%Y/%m/%d")

    # タイトル (A1) - T列までマージ
    ws.merge_cells('A1:T1')
    title_cell = ws['A1']
    title_text = f"テスト進捗サマリー（{team_name}）" if team_name != "ALL" else "テスト進捗サマリー（全体）"
    title_cell.value = title_text
    title_cell.font = TITLE_FONT
    title_cell.fill = TITLE_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25

    # 集計情報（A2）+ 基準日（R2ラベル、S2:T2結合で値）
    ws['A2'] = f"集計期間: {min_date.strftime('%Y/%m/%d')} ～ {max_date.strftime('%Y/%m/%d')} ({len(date_range)}日間)"
    ws['A2'].font = Font(name="游ゴシック", size=10, color="666666")

    ws['R2'] = "基準日:"
    ws['R2'].font = Font(name="游ゴシック", size=11, bold=True)
    ws['R2'].alignment = DATA_ALIGN_RIGHT

    ws.merge_cells('S2:T2')
    ws['S2'] = today_str
    ws['S2'].font = REF_DATE_FONT
    ws['S2'].fill = REF_DATE_FILL
    ws['S2'].alignment = DATA_ALIGN_CENTER

    # ヘッダー行 (row 4)
    # 列構成: A-T（20列）、U列は基準日参照用（非表示扱い）
    header_row = 4
    summary_headers = [
        "日付", "曜日", "営業日",
        "実施_予定", "実施_実績", "実施_消化率", "実施_予定累計", "実施_実績累計", "実施_累計消化率", "実施_判定",
        "検証_予定", "検証_実績", "検証_消化率", "検証_予定累計", "検証_実績累計", "検証_累計消化率", "検証_判定",
        "合計_予定", "合計_実績", "基準日",
    ]

    for col, header in enumerate(summary_headers, 1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    # 基準日の値はU2に配置（数式参照用、表外）
    ws['U2'] = today_str
    ws['U2'].font = Font(name="游ゴシック", size=10)
    ws['U2'].alignment = DATA_ALIGN_CENTER

    # データ行
    data_start_row = header_row + 1
    detail_last_row = detail_start_row + total_record_count - 1

    # チームフィルタ用のCOUNTIFS条件
    if team_name == "ALL":
        # ALLの場合はチーム条件なし
        team_condition = ""
    else:
        # チーム別の場合は明細のD列（チーム名）でフィルタ
        team_condition = f',明細!$D${detail_start_row}:$D${detail_last_row},"{team_name}"'

    for i, date_obj in enumerate(date_range):
        row = data_start_row + i
        date_str = date_obj.strftime("%Y/%m/%d")

        # 曜日・営業日は数式で計算
        # B列: 曜日 = CHOOSE(WEEKDAY(A列,2),"月","火","水","木","金","土","日")
        weekday_formula = f'=CHOOSE(WEEKDAY(A{row},2),"月","火","水","木","金","土","日")'
        # C列: 営業日 = IF(OR(WEEKDAY(A列,2)>=6, COUNTIF(祝日マスタ!$A:$A,A列)>0),"非営業日","営業日")
        business_formula = f'=IF(OR(WEEKDAY(A{row},2)>=6,COUNTIF(祝日マスタ!$A:$A,A{row})>0),"非営業日","営業日")'

        # 明細シートを参照するCOUNTIFS関数
        # 実施者_予定 (F列=6)
        jisshi_yotei = f'=COUNTIFS(明細!$F${detail_start_row}:$F${detail_last_row},A{row}{team_condition})'
        # 実施者_実績 (G列=7)
        jisshi_jisseki = f'=COUNTIFS(明細!$G${detail_start_row}:$G${detail_last_row},A{row}{team_condition})'
        # 実施者_消化率
        jisshi_rate = f'=IF(D{row}=0,0,E{row}/D{row})'
        # 実施者_予定累計
        jisshi_yotei_cum = f'=SUM($D${data_start_row}:D{row})'
        # 実施者_実績累計
        jisshi_jisseki_cum = f'=SUM($E${data_start_row}:E{row})'
        # 実施者_累計消化率
        jisshi_cum_rate = f'=IF(G{row}=0,0,H{row}/G{row})'
        # 実施_判定（累計実績>=累計予定なら順調、累計消化率100%なら完了、日付>基準日なら予定）
        jisshi_status = f'=IF(A{row}>$U$2,"予定",IF(I{row}>=1,"完了",IF(H{row}>=G{row},"順調","遅延")))'

        # 検証者_予定 (I列=9 in 明細)
        kensho_yotei = f'=COUNTIFS(明細!$I${detail_start_row}:$I${detail_last_row},A{row}{team_condition})'
        # 検証者_実績 (J列=10 in 明細)
        kensho_jisseki = f'=COUNTIFS(明細!$J${detail_start_row}:$J${detail_last_row},A{row}{team_condition})'
        # 検証者_消化率
        kensho_rate = f'=IF(K{row}=0,0,L{row}/K{row})'
        # 検証者_予定累計
        kensho_yotei_cum = f'=SUM($K${data_start_row}:K{row})'
        # 検証者_実績累計
        kensho_jisseki_cum = f'=SUM($L${data_start_row}:L{row})'
        # 検証者_累計消化率
        kensho_cum_rate = f'=IF(N{row}=0,0,O{row}/N{row})'
        # 検証_判定
        kensho_status = f'=IF(A{row}>$U$2,"予定",IF(P{row}>=1,"完了",IF(O{row}>=N{row},"順調","遅延")))'

        # 合計
        total_yotei = f'=D{row}+K{row}'
        total_jisseki = f'=E{row}+L{row}'

        # 基準日マーク（T列）: 今日の日付と一致すれば★
        ref_mark = f'=IF(A{row}=$U$2,"★","")'

        values = [
            date_str,           # A: 日付
            weekday_formula,    # B: 曜日（数式）
            business_formula,   # C: 営業日（数式）
            jisshi_yotei,       # D: 実施_予定
            jisshi_jisseki,     # E: 実施_実績
            jisshi_rate,        # F: 実施_消化率
            jisshi_yotei_cum,   # G: 実施_予定累計
            jisshi_jisseki_cum, # H: 実施_実績累計
            jisshi_cum_rate,    # I: 実施_累計消化率
            jisshi_status,      # J: 実施_判定
            kensho_yotei,       # K: 検証_予定
            kensho_jisseki,     # L: 検証_実績
            kensho_rate,        # M: 検証_消化率
            kensho_yotei_cum,   # N: 検証_予定累計
            kensho_jisseki_cum, # O: 検証_実績累計
            kensho_cum_rate,    # P: 検証_累計消化率
            kensho_status,      # Q: 検証_判定
            total_yotei,        # R: 合計_予定
            total_jisseki,      # S: 合計_実績
            ref_mark,           # T: 基準日マーク
        ]

        for col, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = DATA_FONT
            cell.border = THIN_BORDER

            # パーセント列
            if col in (6, 9, 13, 16):  # F, I, M, P
                cell.number_format = "0.0%"
                cell.alignment = DATA_ALIGN_RIGHT
            elif col == 1:
                cell.alignment = DATA_ALIGN_CENTER
            elif col in (2, 3, 10, 17, 20):  # B, C, J, Q, T
                cell.alignment = DATA_ALIGN_CENTER
            else:
                cell.alignment = DATA_ALIGN_CENTER


    # 合計行
    last_data_row = data_start_row + len(date_range) - 1
    total_row = last_data_row + 1

    ws.cell(row=total_row, column=1, value="合計").font = Font(name="游ゴシック", size=11, bold=True)
    ws.cell(row=total_row, column=1).alignment = DATA_ALIGN_CENTER
    ws.cell(row=total_row, column=1).border = THIN_BORDER
    ws.cell(row=total_row, column=1).fill = TOTAL_FILL

    # 合計列の設定
    sum_cols = [4, 5, 7, 8, 11, 12, 14, 15, 18, 19]  # D,E,G,H,K,L,N,O,R,S
    rate_cols = {6: (5, 4), 9: (8, 7), 13: (12, 11), 16: (15, 14)}  # 消化率列: (分子, 分母)

    for col in range(2, 21):
        cell = ws.cell(row=total_row, column=col)
        cell.border = THIN_BORDER
        cell.fill = TOTAL_FILL
        cell.font = Font(name="游ゴシック", size=11, bold=True)

        if col in sum_cols:
            cl = get_column_letter(col)
            cell.value = f"=SUM({cl}{data_start_row}:{cl}{last_data_row})"
            cell.alignment = DATA_ALIGN_CENTER
        elif col in rate_cols:
            num_col, den_col = rate_cols[col]
            nc = get_column_letter(num_col)
            dc = get_column_letter(den_col)
            cell.value = f"=IF({dc}{total_row}=0,0,{nc}{total_row}/{dc}{total_row})"
            cell.number_format = "0.0%"
            cell.alignment = DATA_ALIGN_RIGHT
        elif col in (2, 3, 10, 17, 20):
            cell.value = ""

    # テーブル作成（テーブル名はシートごとにユニークに）
    if date_range:
        table_name = f"サマリー_{team_name.replace(' ', '_')}"
        table_ref = f"A{header_row}:T{last_data_row}"
        table = Table(displayName=table_name, ref=table_ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    # 条件付き書式: 消化率列（0-100%のカラースケール）
    rate_columns = ['F', 'I', 'M', 'P']
    for col_letter in rate_columns:
        range_str = f"{col_letter}{data_start_row}:{col_letter}{last_data_row}"
        ws.conditional_formatting.add(
            range_str,
            ColorScaleRule(
                start_type='num', start_value=0, start_color='FFC7CE',      # 赤
                mid_type='num', mid_value=0.5, mid_color='FFEB9C',          # 黄
                end_type='num', end_value=1, end_color='C6EFCE'             # 緑
            )
        )

    # 条件付き書式: 実施_判定列（J列）、検証_判定列（Q列）
    for status_col in ['J', 'Q']:
        # 完了 = 緑
        ws.conditional_formatting.add(
            f"{status_col}{data_start_row}:{status_col}{last_data_row}",
            FormulaRule(
                formula=[f'${status_col}{data_start_row}="完了"'],
                fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                font=Font(color="006100", bold=True)
            )
        )
        # 順調 = 青
        ws.conditional_formatting.add(
            f"{status_col}{data_start_row}:{status_col}{last_data_row}",
            FormulaRule(
                formula=[f'${status_col}{data_start_row}="順調"'],
                fill=PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
                font=Font(color="1F4E79", bold=True)
            )
        )
        # 遅延 = 赤
        ws.conditional_formatting.add(
            f"{status_col}{data_start_row}:{status_col}{last_data_row}",
            FormulaRule(
                formula=[f'${status_col}{data_start_row}="遅延"'],
                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                font=Font(color="9C0006", bold=True)
            )
        )
        # 予定 = グレー
        ws.conditional_formatting.add(
            f"{status_col}{data_start_row}:{status_col}{last_data_row}",
            FormulaRule(
                formula=[f'${status_col}{data_start_row}="予定"'],
                fill=PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
                font=Font(color="808080")
            )
        )

    # 条件付き書式: 非営業日は行全体をグレーに（C列="非営業日"）
    ws.conditional_formatting.add(
        f"A{data_start_row}:T{last_data_row}",
        FormulaRule(
            formula=[f'$C{data_start_row}="非営業日"'],
            fill=COMPLETED_ROW_FILL
        )
    )

    # 基準日マーク列（T列）のスタイル
    ws.conditional_formatting.add(
        f"T{data_start_row}:T{last_data_row}",
        FormulaRule(
            formula=[f'$T{data_start_row}="★"'],
            fill=PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"),
            font=Font(color="000000", bold=True, size=14)
        )
    )

    # 列幅設定
    summary_widths = [12, 5, 10, 9, 9, 10, 11, 11, 11, 8, 9, 9, 10, 11, 11, 11, 8, 9, 9, 6]
    for i, w in enumerate(summary_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # U列（基準日参照用、非表示）
    ws.column_dimensions['U'].hidden = True

    ws.freeze_panes = f"D{data_start_row}"


# ===================================================================
#  メイン処理
# ===================================================================

def main():
    print("=" * 60)
    print("  テスト予定・実績 集計スクリプト v4")
    print("=" * 60)

    # --- ウィザードUI実行 ---
    config = run_wizard()

    if config is None:
        print("\n  キャンセルされました。")
        sys.exit(0)

    folder_path = config["folder_path"]
    output_path = config["output_path"]
    include_subfolders = config["include_subfolders"]

    # キャッシュファイルのパス（出力ファイルと同じディレクトリ）
    cache_dir = os.path.dirname(output_path)
    cache_file = os.path.join(cache_dir, ".test_collector_cache.json")

    print(f"\n  対象フォルダ: {folder_path}")
    print(f"  サブフォルダ: {'含める' if include_subfolders else '含めない'}")
    print(f"  出力先:       {output_path}")
    print(f"  対象シート:   {SHEET_PREFIX}* で始まるシート\n")

    records = collect_data(folder_path, cache_file, include_subfolders)

    if not records:
        msg = (
            "データが見つかりませんでした。\n"
            "・フォルダパスを確認してください\n"
            f"・シート名が「{SHEET_PREFIX}」で始まるか確認してください\n"
            "・C列にテストIDが入っているか確認してください"
        )
        print(f"\n  ⚠ {msg}")
        root_err = tk.Tk()
        root_err.withdraw()
        root_err.attributes("-topmost", True)
        messagebox.showwarning("データなし", msg)
        root_err.destroy()
        sys.exit(1)

    write_excel(records, output_path)

    print("\n" + "=" * 60)

    # 完了メッセージ
    team_counts = defaultdict(int)
    for rec in records:
        team_counts[rec["チーム名"]] += 1

    team_info = "\n".join([f"  - {team}: {count}件" for team, count in sorted(team_counts.items())])

    root_done = tk.Tk()
    root_done.withdraw()
    root_done.attributes("-topmost", True)
    messagebox.showinfo(
        "完了",
        f"集計が完了しました！\n\n"
        f"出力先:\n{output_path}\n\n"
        f"明細: {len(records)}件\n\n"
        f"チーム別内訳:\n{team_info}",
    )
    root_done.destroy()


if __name__ == "__main__":
    main()
