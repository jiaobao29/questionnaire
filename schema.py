# schema.py

# ── 檔案路徑設定 ──
OUTPUT_PATH   = "questionnaire_filled.xlsx"
PROGRESS_FILE = "progress.json"

# ── 問卷格式設定 ──
QUESTIONS_COUNT = 4
OPTIONS_RANGE = "12345"

# ── 欄位映射對應 (Column Mapping) ──
# 以 0-indexed 為基底，寫入 Excel 時需 +1
QUESTION_COL_START = [1, 6, 11, 16]
TEXT_COL_1 = 21   # Excel col 22 (V) — immediately after Q4's 5 options
TEXT_COL_2 = 22   # Excel col 23 (W)

# 資料開始寫入的列數
DATA_START_ROW = 5
