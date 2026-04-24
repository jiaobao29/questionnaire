# app.py
import streamlit as st
import logic
from schema import OPTIONS_RANGE, DATA_START_ROW

st.set_page_config(page_title="問卷快速輸入系統", page_icon="🚀", layout="wide")

def main():
    st.title("🚀 問卷快速輸入系統")
    
    # ── SIDEBAR: 設定、下載與重置 ──
    with st.sidebar:
        st.header("⚙️ 系統設定")
        q_count = st.number_input(
            "選擇題數量 (N)", 
            min_value=1, 
            max_value=50, 
            value=st.session_state.get('q_count', 4),
            step=1,
            help="動態調整問卷的選擇題數量"
        )
        
        # 如果 q_count 改變，我們需要通知使用者重置或自動處理
        if 'q_count' in st.session_state and st.session_state.q_count != q_count:
            st.warning("⚠️ 檢測到問題數量變動！")
            if st.button("🔄 套用變動並重置 Excel 格式"):
                logic.reset_data()
                st.session_state.q_count = q_count
                st.session_state.current_row = DATA_START_ROW
                st.session_state.session_count = 0
                st.rerun()
        else:
            st.session_state.q_count = q_count

        st.divider()
        st.header("📥 資料下載與管理")
        
        # 初始化系統 (使用當前的 q_count)
        is_ok = logic.initialize_system(st.session_state.q_count)
        if not is_ok:
            st.error(f"❌ 現有的 Excel 格式與問題數量 ({st.session_state.q_count}) 不符！")
            if st.button("🗑️ 重置 Excel 以符合新格式"):
                logic.reset_data()
                st.rerun()
            return # 停止渲染其餘部分

        # 此時讀取的檔案二進位資料
        excel_bytes = logic.get_excel_download_bytes()
        if excel_bytes:
            st.download_button(
                label="📥 下載完成的問卷 Excel",
                data=excel_bytes,
                file_name=f"questionnaire_{st.session_state.q_count}q_filled.xlsx",
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

    # 1. Session State 初始化 (確保 current_row 就緒)
    if 'current_row' not in st.session_state:
        st.session_state.current_row = logic.load_progress()
    if 'session_count' not in st.session_state:
        st.session_state.session_count = 0

    # 2. 載入 Excel
    wb = logic.get_workbook()

    # 3. 計算統計 (傳入動態 q_count)
    stats_df = logic.update_and_get_stats(wb, st.session_state.current_row, st.session_state.q_count)

    # 目前是第幾份
    seq = st.session_state.current_row - DATA_START_ROW + 1
    
    # ── MAIN UI 佈局 ──
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.info(f"📈 目前進度: **第 {seq} 份** | 問題數量: **{st.session_state.q_count}** | 本次輸入: {st.session_state.session_count} 份")
        
        with st.form("entry_form", clear_on_submit=True):
            st.subheader("📝 資料輸入")
            choices_input = st.text_input(
                f"🔢 選擇題 ({st.session_state.q_count}位數, 1~5)", 
                max_chars=st.session_state.q_count,
                placeholder=f"例如: {'1' * st.session_state.q_count}"
            ).strip().lower()
            
            t1 = st.text_input("💬 反映與建議", placeholder="跳過請留空")
            t2 = st.text_input("💬 其他建議", placeholder="跳過請留空")
            
            st.markdown("*提示：輸入完畢後按 `Enter` 即可快速送出*")
            submitted = st.form_submit_button("🚀 送出 (Submit)", use_container_width=True)

        if submitted:
            # 驗證邏輯
            if not choices_input:
                st.error("⚠️ 選擇題不能為空！")
            elif len(choices_input) != st.session_state.q_count or not all(c in OPTIONS_RANGE for c in choices_input):
                st.error(f"⚠️ 格式錯誤！請輸入 {st.session_state.q_count} 個 1~5 的數字")
            else:
                choices = [int(c) for c in choices_input]
                
                # 呼叫邏輯層將原始資料寫入 Excel
                logic.save_survey_entry(wb, st.session_state.current_row, choices, t1, t2, st.session_state.q_count)
                
                # 推進前端 State 狀態
                st.session_state.current_row += 1
                st.session_state.session_count += 1
                
                st.success(f"✅ 成功儲存第 {seq} 份！")
                st.rerun()

    with col2:
        st.subheader("📊 即時統計報表")
        if stats_df is not None:
            st.dataframe(stats_df, hide_index=True, use_container_width=True)
        else:
            st.write("⚠️ 目前尚無資料可統計。開始輸入以查看圖表！")

if __name__ == "__main__":
    main()
