# schema.py

# ── 檔案路徑設定 ──
OUTPUT_PATH   = "questionnaire_filled.xlsx"
PROGRESS_FILE = "progress.json"

# ── 問卷格式設定 ──
# QUESTIONS_COUNT = 4  # 移至 app.py 由使用者動態設定
OPTIONS_RANGE = "12345"

# ── 欄位映射對應 (Column Mapping) ──
def get_config(q_count: int):
    """
    根據問題數量 N 動態計算欄位位置
    N 個問題，每個問題佔 5 欄 (非常滿意~非常不滿意)
    """
    # 每個問題的起始欄位索引 (0-indexed)
    # Q1: 1, Q2: 6, Q3: 11, ...
    question_starts = list(range(1, (q_count * 5) + 1, 5))
    
    # 文字題 1 緊接在最後一個問題的 5 個選項之後
    text_col_1 = (q_count * 5) + 1
    # 文字題 2
    text_col_2 = (q_count * 5) + 2
    
    return {
        "question_starts": question_starts,
        "text_col_1": text_col_1,
        "text_col_2": text_col_2
    }

# 資料開始寫入的列數
DATA_START_ROW = 5
