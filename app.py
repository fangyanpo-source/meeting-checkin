import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import json  # 雲端讀取金鑰必備

# ==========================================
# 1. 頁面與基本設定
# ==========================================
st.set_page_config(page_title="會議報到工作站", page_icon="📋", layout="wide")

# ==========================================
# 2. 連線 Google Sheets (雲端安全版)
# ==========================================
@st.cache_resource
def init_gsheets_client():
    """初始化並回傳 Google Sheets Spreadsheet 物件（不含特定工作表）"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        secret_dict = json.loads(st.secrets["gcp_json"])
        creds = Credentials.from_service_account_info(secret_dict, scopes=scopes)
        client = gspread.authorize(creds)
        # ⚠️ 【重要修改處】：請把下面這串亂碼換成你真實的「試算表 ID」
        spreadsheet = client.open_by_key("1Ezc7IVTQJF76pSCrZEsF2n9vW_dBEPDyYrsx2-jwZOI")
        return spreadsheet
    except Exception as e:
        st.error(f"資料庫連線失敗，請檢查金鑰設定。錯誤訊息：{e}")
        st.stop()

def get_sheet(spreadsheet, session_name):
    """依時段名稱取得對應工作表，若不存在則自動建立"""
    try:
        return spreadsheet.worksheet(session_name)
    except gspread.exceptions.WorksheetNotFound:
        # 自動建立工作表並加上標題列
        ws = spreadsheet.add_worksheet(title=session_name, rows=200, cols=10)
        ws.append_row(["姓名", "單位", "報到狀態", "報到時間"])
        return ws

spreadsheet = init_gsheets_client()

# ==========================================
# 2b. 側邊欄：時段選擇
# ==========================================
SESSIONS = {
    "🌅 上午場": "上午",
    "🌇 下午場": "下午",
}

with st.sidebar:
    st.header("⏰ 時段選擇")
    selected_label = st.radio(
        "請選擇要操作的時段：",
        options=list(SESSIONS.keys()),
        index=0,
    )
    selected_session = SESSIONS[selected_label]

    st.divider()
    st.caption("每個時段對應 Google Sheets 中獨立的工作表分頁。")

# 取得當前時段的工作表物件
sheet = get_sheet(spreadsheet, selected_session)

# ==========================================
# 3. 讀取與初始化資料
# ==========================================
def load_data(ws):
    """從指定工作表讀取資料"""
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if not df.empty and '報到狀態' in df.columns:
        df['報到狀態'] = df['報到狀態'].astype(str).str.upper() == 'TRUE'
    return df

# 切換時段時重新載入資料（以 session_name 為 key 區分）
session_key = f"attendees_{selected_session}"
if session_key not in st.session_state:
    st.session_state[session_key] = load_data(sheet)

# ==========================================
# 4. 頂部儀表板：報到統計
# ==========================================
st.title(f"📋 現場會議報到　｜　{selected_label}")

# 計算統計數據
df = st.session_state[session_key]
if not df.empty:
    total_people = len(df)
    checked_in_people = df['報到狀態'].sum()
    check_in_rate = (checked_in_people / total_people) * 100 if total_people > 0 else 0
else:
    total_people = checked_in_people = check_in_rate = 0

col_stat1, col_stat2, col_stat3 = st.columns(3)
col_stat1.metric("總人數", f"{total_people} 人")
col_stat2.metric("已報到", f"{checked_in_people} 人")
col_stat3.metric("報到率", f"{check_in_rate:.1f} %")

st.divider()

# ==========================================
# 5. 搜尋與篩選區
# ==========================================
search_term = st.text_input("🔍 快速搜尋：請輸入姓名或單位", placeholder="例如：資訊部 或 王大明")
tab1, tab2, tab3 = st.tabs(["全部名單", "⏳ 尚未報到", "✅ 已報到"])

if search_term and not df.empty:
    df = df[df['姓名'].astype(str).str.contains(search_term) | df['單位'].astype(str).str.contains(search_term)]

# ==========================================
# 6. 渲染名單列表函式 (含雲端寫入邏輯)
# ==========================================
def render_list(data_frame, prefix):
    if data_frame.empty:
        st.info("沒有找到符合條件的人員，或試算表目前為空。")
        return

    for index, row in data_frame.iterrows():
        col_info, col_action = st.columns([6, 4], vertical_alignment="center")

        with col_info:
            st.markdown(f"**{row['姓名']}** |  {row['單位']}")
            if row['報到狀態']:
                st.caption(f"🕒 報到時間：{row['報到時間']}")
            else:
                st.caption("🔴 未報到")

        with col_action:
            if not row['報到狀態']:
                if st.button("👉 點此簽到", key=f"checkin_{prefix}_{selected_session}_{index}", use_container_width=True, type="primary"):
                    current_time = datetime.now().strftime("%H:%M:%S")

                    # --- 1. 更新 Google Sheets 雲端資料 ---
                    # 試算表有標題列(佔1列)，且 Pandas index 從 0 開始，所以雲端實際列數 = index + 2
                    try:
                        sheet.update_cell(index + 2, 3, "TRUE")
                        sheet.update_cell(index + 2, 4, current_time)

                        # --- 2. 更新本地端 Session State ---
                        st.session_state[session_key].at[index, '報到狀態'] = True
                        st.session_state[session_key].at[index, '報到時間'] = current_time

                        # 顯示成功提示並重新整理畫面
                        st.toast(f"✅ {row['姓名']} 報到成功！已同步至雲端", icon="✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"寫入雲端失敗，請確認網路連線或金鑰權限。錯誤代碼：{e}")
            else:
                st.button("已完成", key=f"done_{prefix}_{selected_session}_{index}", disabled=True, use_container_width=True)

        st.divider()

# ==========================================
# 7. 在不同頁籤中顯示對應資料
# ==========================================
if not df.empty:
    with tab1:
        render_list(df, prefix="all")
    with tab2:
        render_list(df[df['報到狀態'] == False], prefix="pending")
    with tab3:
        render_list(df[df['報到狀態'] == True], prefix="completed")