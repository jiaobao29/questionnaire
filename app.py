# app.py
import streamlit as st
import logic
from schema import QUESTIONS_COUNT, OPTIONS_RANGE, DATA_START_ROW

st.set_page_config(page_title="問卷快速輸入系統", page_icon="🚀", layout="wide")

def main():
    st.title("🚀 問卷快速輸入系統")
    
    # 1. 系統初始化 (確保 Excel 檔案就緒)
    logic.initialize_system()

    # 2. Session State 初始化
    if 'current_row' not in st.session_state:
        st.session_state.current_row = logic.load_progress()
    if 'session_count' not in st.session_state:
        st.session_state.session_count = 0

    # 載入 Excel (為了畫面上的統計計算與資料輸入)
    wb = logic.get_workbook()

    # ── SIDEBAR: 下載與重置 ──
    with st.sidebar:
        st.header("📥 資料下載與管理")
        st.markdown("輸入完畢後，請點擊下方按鈕將結果下載回您的電腦。")
        
        excel_bytes = logic.get_excel_download_bytes()
        if excel_bytes:
            st.download_button(
                label="📥 下載完成的問卷 Excel",
                data=excel_bytes,
                file_name="questionnaire_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        st.divider()
        st.markdown("⚠️ **重新開始新的一批**")
        if st.button("🗑️ 刪除全部資料並重新開始", use_container_width=True):
            logic.reset_data()
            st.session_state.current_row = DATA_START_ROW
            st.session_state.session_count = 0
            st.rerun()

    # 目前是第幾份
    seq = st.session_state.current_row - DATA_START_ROW + 1
    
    # ── MAIN UI 佈局 ──
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.info(f"📈 目前進度: **第 {seq} 份** (Excel 第 {st.session_state.current_row} 列) | 本次輸入: {st.session_state.session_count} 份")
        
        with st.form("entry_form", clear_on_submit=True):
            st.subheader("📝 資料輸入")
            choices_input = st.text_input(
                f"🔢 選擇題 ({QUESTIONS_COUNT}位數, 1~5)", 
                max_chars=QUESTIONS_COUNT,
                placeholder="例如: 1121"
            ).strip().lower()
            
            t1 = st.text_input("💬 反映與建議", placeholder="跳過請留空")
            t2 = st.text_input("💬 其他建議", placeholder="跳過請留空")
            
            st.markdown("*提示：輸入完畢後按 `Enter` 即可快速送出*")
            submitted = st.form_submit_button("🚀 送出 (Submit)", use_container_width=True)

        if submitted:
            # 驗證邏輯
            if not choices_input:
                st.error("⚠️ 選擇題不能為空！")
            elif len(choices_input) != QUESTIONS_COUNT or not all(c in OPTIONS_RANGE for c in choices_input):
                st.error(f"⚠️ 格式錯誤！請輸入 {QUESTIONS_COUNT} 個 1~5 的數字 (例如: 1121)")
            else:
                choices = [int(c) for c in choices_input]
                
                # 呼叫邏輯層將資料寫入 Excel 並更新 JSON 進度
                logic.save_survey_entry(wb, st.session_state.current_row, choices, t1, t2)
                
                # 推進前端 State 狀態
                st.session_state.current_row += 1
                st.session_state.session_count += 1
                
                st.success(f"✅ 成功儲存第 {seq} 份！")
                st.rerun()

    with col2:
        st.subheader("📊 即時統計報表")
        # 直接拿 DataFrame 回來呈現，不接觸 openpyxl 邏輯
        stats_df = logic.update_and_get_stats(wb, st.session_state.current_row)
        
        if stats_df is not None:
            st.dataframe(stats_df, hide_index=True, use_container_width=True)
        else:
            st.write("⚠️ 目前尚無資料可統計。開始輸入以查看圖表！")

if __name__ == "__main__":
    main()
