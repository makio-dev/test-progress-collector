"""
テスト進捗集計ツール - 検証用データ生成スクリプト

6ヶ月のプロジェクト（2025/12〜2026/05）を想定した
各種機能の確認が網羅できるテストデータを生成します。

生成されるデータ:
- 4チーム（オンライン、バッチ、基盤、運用）+ その他
- 各チームに複数のテストケースファイル
- 様々な進捗状況（完了、遅延、予定、未着手）
- 日付パターン（過去、当日、未来）
- サブフォルダ構造
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import os
from datetime import datetime, timedelta
import random

# 設定
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "input")
SHEET_PREFIX = "ITB"
DATA_START_ROW = 19
COL_TEST_ID = 3  # C列
COL_JISSHI_YOTEI = 17  # Q列
COL_JISSHI_JISSEKI = 18  # R列
COL_KENSHO_YOTEI = 19  # S列
COL_KENSHO_JISSEKI = 20  # T列

# プロジェクト期間
PROJECT_START = datetime(2025, 12, 1)
PROJECT_END = datetime(2026, 5, 31)
TODAY = datetime(2026, 3, 5)  # 基準日

# チーム設定
TEAMS = {
    "オンライン": {
        "pattern": "-O-",
        "files": [
            {"name": "ITB-O-001_ログイン認証", "sheets": ["ITB001_ログイン", "ITB002_ログアウト", "ITB003_パスワード変更"], "cases_per_sheet": 15},
            {"name": "ITB-O-002_ユーザー管理", "sheets": ["ITB001_ユーザー登録", "ITB002_ユーザー編集", "ITB003_ユーザー削除"], "cases_per_sheet": 12},
            {"name": "ITB-O-003_検索機能", "sheets": ["ITB001_商品検索", "ITB002_注文検索"], "cases_per_sheet": 20},
            {"name": "ITB-O-004_カート機能", "sheets": ["ITB001_カート追加", "ITB002_カート編集", "ITB003_カート削除"], "cases_per_sheet": 10},
        ],
        "subdir": None,
    },
    "バッチ": {
        "pattern": "-B-",
        "files": [
            {"name": "ITB-B-001_日次バッチ", "sheets": ["ITB001_売上集計", "ITB002_在庫更新", "ITB003_レポート生成"], "cases_per_sheet": 8},
            {"name": "ITB-B-002_月次バッチ", "sheets": ["ITB001_月次締め", "ITB002_請求書発行"], "cases_per_sheet": 15},
            {"name": "ITB-B-003_データ連携", "sheets": ["ITB001_外部システム連携", "ITB002_データ同期"], "cases_per_sheet": 10},
        ],
        "subdir": None,
    },
    "基盤": {
        "pattern": "-I-",
        "files": [
            {"name": "ITB-I-001_認証基盤", "sheets": ["ITB001_SSO", "ITB002_トークン管理", "ITB003_権限制御"], "cases_per_sheet": 12},
            {"name": "ITB-I-002_ログ基盤", "sheets": ["ITB001_アクセスログ", "ITB002_エラーログ", "ITB003_監査ログ"], "cases_per_sheet": 8},
        ],
        "subdir": "基盤チーム",
    },
    "運用": {
        "pattern": "-U-",
        "files": [
            {"name": "ITB-U-001_監視機能", "sheets": ["ITB001_死活監視", "ITB002_性能監視", "ITB003_アラート通知"], "cases_per_sheet": 10},
            {"name": "ITB-U-002_運用ツール", "sheets": ["ITB001_バックアップ", "ITB002_リストア", "ITB003_メンテナンス"], "cases_per_sheet": 6},
        ],
        "subdir": "運用チーム",
    },
    "その他": {
        "pattern": "",  # パターンなし
        "files": [
            {"name": "ITB_共通機能テスト", "sheets": ["ITB001_帳票出力", "ITB002_メール送信"], "cases_per_sheet": 8},
            {"name": "ITB_外部連携テスト", "sheets": ["ITB001_API連携"], "cases_per_sheet": 5},
        ],
        "subdir": "その他",
    },
}

# 進捗パターン
def generate_progress_pattern(test_id_num, total_cases, today):
    """テストケースの進捗パターンを生成"""
    progress_ratio = test_id_num / total_cases

    # プロジェクト進捗に基づく予定日を計算
    project_days = (PROJECT_END - PROJECT_START).days
    base_offset = int(project_days * progress_ratio)
    jisshi_yotei = PROJECT_START + timedelta(days=base_offset + random.randint(-5, 5))
    kensho_yotei = jisshi_yotei + timedelta(days=random.randint(1, 5))

    # 進捗状況を決定
    jisshi_jisseki = None
    kensho_jisseki = None

    if jisshi_yotei <= today:
        # 予定日が過去の場合
        if random.random() < 0.85:  # 85%は実施完了
            jisshi_jisseki = jisshi_yotei + timedelta(days=random.randint(-2, 3))
            if kensho_yotei <= today:
                if random.random() < 0.80:  # 80%は検証完了
                    kensho_jisseki = kensho_yotei + timedelta(days=random.randint(-1, 2))
                # 20%は検証遅延
            # 検証予定が未来の場合は検証未実施
        # 15%は実施遅延（実績なし）
    # 予定が未来の場合は未着手

    return jisshi_yotei, jisshi_jisseki, kensho_yotei, kensho_jisseki


def create_test_file(filepath, sheets_config, cases_per_sheet):
    """テストケースファイルを作成"""
    wb = openpyxl.Workbook()

    # 最初のシートを削除用にマーク
    first_sheet = True

    for sheet_name in sheets_config:
        if first_sheet:
            ws = wb.active
            ws.title = sheet_name
            first_sheet = False
        else:
            ws = wb.create_sheet(sheet_name)

        # ヘッダー行（18行目）
        header_row = 18
        headers = {
            COL_TEST_ID: "テストID",
            COL_JISSHI_YOTEI: "実施予定",
            COL_JISSHI_JISSEKI: "実施実績",
            COL_KENSHO_YOTEI: "検証予定",
            COL_KENSHO_JISSEKI: "検証実績",
        }

        for col, header in headers.items():
            cell = ws.cell(row=header_row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        # テストケースデータ
        for i in range(cases_per_sheet):
            row = DATA_START_ROW + i
            test_id = f"{sheet_name}-{i+1:03d}"

            jisshi_yotei, jisshi_jisseki, kensho_yotei, kensho_jisseki = generate_progress_pattern(
                i + 1, cases_per_sheet, TODAY
            )

            ws.cell(row=row, column=COL_TEST_ID, value=test_id)
            ws.cell(row=row, column=COL_JISSHI_YOTEI, value=jisshi_yotei)
            ws.cell(row=row, column=COL_JISSHI_JISSEKI, value=jisshi_jisseki)
            ws.cell(row=row, column=COL_KENSHO_YOTEI, value=kensho_yotei)
            ws.cell(row=row, column=COL_KENSHO_JISSEKI, value=kensho_jisseki)

            # 日付書式
            for col in [COL_JISSHI_YOTEI, COL_JISSHI_JISSEKI, COL_KENSHO_YOTEI, COL_KENSHO_JISSEKI]:
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    cell.number_format = "YYYY/MM/DD"

    wb.save(filepath)
    print(f"  ✅ {os.path.basename(filepath)} ({len(sheets_config)}シート, 各{cases_per_sheet}件)")


def main():
    """メイン処理"""
    print("=" * 60)
    print("  テスト進捗集計ツール - 検証用データ生成")
    print("=" * 60)
    print(f"\n  プロジェクト期間: {PROJECT_START.strftime('%Y/%m/%d')} 〜 {PROJECT_END.strftime('%Y/%m/%d')}")
    print(f"  基準日: {TODAY.strftime('%Y/%m/%d')}")
    print(f"  出力先: {OUTPUT_DIR}\n")

    # 既存ファイルを削除
    if os.path.exists(OUTPUT_DIR):
        for root, dirs, files in os.walk(OUTPUT_DIR, topdown=False):
            for name in files:
                if name.endswith(('.xlsx', '.xlsm')) and not name.startswith('~$'):
                    os.remove(os.path.join(root, name))
            for name in dirs:
                try:
                    os.rmdir(os.path.join(root, name))
                except OSError:
                    pass

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    total_files = 0
    total_sheets = 0
    total_cases = 0

    for team_name, team_config in TEAMS.items():
        print(f"\n[{team_name}チーム]")

        # サブディレクトリ
        if team_config["subdir"]:
            team_dir = os.path.join(OUTPUT_DIR, team_config["subdir"])
            os.makedirs(team_dir, exist_ok=True)
        else:
            team_dir = OUTPUT_DIR

        for file_config in team_config["files"]:
            filename = f"{file_config['name']}.xlsx"
            filepath = os.path.join(team_dir, filename)

            create_test_file(
                filepath,
                file_config["sheets"],
                file_config["cases_per_sheet"]
            )

            total_files += 1
            total_sheets += len(file_config["sheets"])
            total_cases += len(file_config["sheets"]) * file_config["cases_per_sheet"]

    # 特殊ケース: xlsmファイル
    print("\n[特殊ケース]")
    xlsm_path = os.path.join(OUTPUT_DIR, "ITB-O-005_マクロ付きテスト.xlsm")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ITB001_マクロテスト"

    # ヘッダー
    for col, header in [(COL_TEST_ID, "テストID"), (COL_JISSHI_YOTEI, "実施予定"),
                        (COL_JISSHI_JISSEKI, "実施実績"), (COL_KENSHO_YOTEI, "検証予定"),
                        (COL_KENSHO_JISSEKI, "検証実績")]:
        ws.cell(row=18, column=col, value=header).font = Font(bold=True)

    # データ
    for i in range(5):
        row = DATA_START_ROW + i
        ws.cell(row=row, column=COL_TEST_ID, value=f"ITB001_マクロテスト-{i+1:03d}")
        ws.cell(row=row, column=COL_JISSHI_YOTEI, value=TODAY - timedelta(days=10-i*2))
        ws.cell(row=row, column=COL_JISSHI_JISSEKI, value=TODAY - timedelta(days=8-i*2) if i < 3 else None)
        ws.cell(row=row, column=COL_KENSHO_YOTEI, value=TODAY - timedelta(days=5-i*2))
        ws.cell(row=row, column=COL_KENSHO_JISSEKI, value=TODAY - timedelta(days=3-i*2) if i < 2 else None)
        for col in [COL_JISSHI_YOTEI, COL_JISSHI_JISSEKI, COL_KENSHO_YOTEI, COL_KENSHO_JISSEKI]:
            cell = ws.cell(row=row, column=col)
            if cell.value:
                cell.number_format = "YYYY/MM/DD"

    wb.save(xlsm_path)
    print(f"  ✅ ITB-O-005_マクロ付きテスト.xlsm (xlsm形式)")
    total_files += 1
    total_sheets += 1
    total_cases += 5

    # 対象外ファイル（ITBで始まらないシート）
    non_target_path = os.path.join(OUTPUT_DIR, "参考資料_設計書.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "設計書"  # ITBで始まらない
    ws['A1'] = "これは参考資料です（集計対象外）"
    wb.save(non_target_path)
    print(f"  ✅ 参考資料_設計書.xlsx (対象外ファイル)")

    print("\n" + "=" * 60)
    print(f"  生成完了!")
    print(f"  - ファイル数: {total_files}")
    print(f"  - シート数: {total_sheets}")
    print(f"  - テストケース数: {total_cases}")
    print("=" * 60)


if __name__ == "__main__":
    main()
