@echo off
chcp 65001 > nul
setlocal
REM ============================================================
REM  テスト進捗集計 → PB曲線 フル生成バッチ
REM    [1/2] aggregate_test_results でダッシュボードExcelを作成
REM    [2/2] generate_pb_curve でPB曲線Excelを作成
REM  ダブルクリックで一気通貫実行できます。
REM ============================================================

REM このバッチがあるフォルダを基準にする
cd /d "%~dp0"

REM ====== 設定（必要に応じて編集） ============================
REM 入力テストケースフォルダ（スペース区切りで複数指定可）
set "INPUT_DIRS=.\input"

REM 出力先
set "OUTPUT_DIR=.\output"
set "DASHBOARD=%OUTPUT_DIR%\dashboard.xlsx"
set "PB_CURVE=%OUTPUT_DIR%\pb_curve.xlsx"

REM 欠陥一覧ファイル（不要なものは空にするか行頭に REM を付けて無効化）
set "DEFECT_ONLINE=.\input\defects\欠陥一覧_オンライン.xlsx"
set "DEFECT_BATCH=.\input\defects\欠陥一覧_バッチ.xlsx"
set "DEFECT_INFRA=.\input\defects\欠陥一覧_基盤.xlsx"
set "DEFECT_OPS=.\input\defects\欠陥一覧_運用.xlsx"

REM PB曲線 B系係数（テストケース数 × 係数 ＝ 欠陥数）
set "B_FINAL_RATE=0.0105"
set "B_LOWER_RATE=0.0035"
set "B_UPPER_RATE=0.0213"

REM 基準日・期間（空なら既定: 前営業日／データの最小・最大日）
set "PIVOT_DATE="
set "START_DATE="
set "END_DATE="

REM 予測倍率・テストケース総数（空なら自動算出）
set "FORECAST_MULT="
set "TOTAL_CASE="
REM ===========================================================

REM Python の決定（.venv があれば優先。EXE運用なら下を書き換え）
set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

REM ---- 欠陥一覧オプションを組み立て（存在するファイルのみ付与） ----
set "DEFECT_OPTS="
if defined DEFECT_ONLINE if exist "%DEFECT_ONLINE%" set DEFECT_OPTS=%DEFECT_OPTS% --defect-online "%DEFECT_ONLINE%"
if defined DEFECT_BATCH  if exist "%DEFECT_BATCH%"  set DEFECT_OPTS=%DEFECT_OPTS% --defect-batch "%DEFECT_BATCH%"
if defined DEFECT_INFRA  if exist "%DEFECT_INFRA%"  set DEFECT_OPTS=%DEFECT_OPTS% --defect-infra "%DEFECT_INFRA%"
if defined DEFECT_OPS    if exist "%DEFECT_OPS%"    set DEFECT_OPTS=%DEFECT_OPTS% --defect-ops "%DEFECT_OPS%"

echo ============================================================
echo [1/2] ダッシュボード集計 -^> %DASHBOARD%
echo ============================================================
"%PY%" aggregate_test_results.py %INPUT_DIRS% -o "%DASHBOARD%"%DEFECT_OPTS%
if errorlevel 1 goto :error

REM ---- PB曲線の任意オプションを組み立て（空でないものだけ付与） ----
set "PB_OPTS="
if defined PIVOT_DATE    set PB_OPTS=%PB_OPTS% --pivot-date %PIVOT_DATE%
if defined START_DATE    set PB_OPTS=%PB_OPTS% --start-date %START_DATE%
if defined END_DATE      set PB_OPTS=%PB_OPTS% --end-date %END_DATE%
if defined FORECAST_MULT set PB_OPTS=%PB_OPTS% --forecast-mult %FORECAST_MULT%
if defined TOTAL_CASE    set PB_OPTS=%PB_OPTS% --total-case %TOTAL_CASE%

echo.
echo ============================================================
echo [2/2] PB曲線生成 -^> %PB_CURVE%
echo ============================================================
"%PY%" generate_pb_curve.py "%DASHBOARD%" -o "%PB_CURVE%" --b-final-rate %B_FINAL_RATE% --b-lower-rate %B_LOWER_RATE% --b-upper-rate %B_UPPER_RATE%%PB_OPTS%
if errorlevel 1 goto :error

echo.
echo 完了しました。PB曲線を開きます: %PB_CURVE%
start "" "%PB_CURVE%"
goto :end

:error
echo.
echo [エラー] 処理に失敗しました。上のログを確認してください。

:end
echo.
pause
endlocal
