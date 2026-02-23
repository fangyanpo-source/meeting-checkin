import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import json

# ==========================================
# 1. 頁面與基本設定 (針對手機螢幕優化)
# ==========================================
st.set_page_config(page_title="會議報到工作站", page_icon="📱", layout="centered")

# 隱藏 Streamlit 預設的右上角選單與底部浮水印 (讓畫面更像原生 App)
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# ==========================================
# 2. 連線 Google Sheets
# ==========================================
@st.cache_resource
def init_gsheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        secret_dict = json.loads(st.secrets["gcp_json"])
        creds = Credentials.from_service_account_info(secret_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        # ⚠️ 【重要】：請替換成你的真實試算表 ID
        sheet = client.open_by_key("1Ezc7IVTQJF76pSCrZEsF2n9vW_dBEPDyYrsx2-jwZOI").sheet1
        return sheet
    except Exception as e:
        st.error(f"資料庫連線失敗：{e}")
        st.stop()

sheet = init_gsheets()

# ==========================================
# 3. 讀取資料
# ==========================================
def load_data():
    records = sheet.get_all_records()
    df = pd.DataFrame(records)
    
    # 確保欄位存在，若舊資料沒有「職稱」則自動補齊
    if '職稱' not in df.columns:
        df['職稱'] = ""
        
    if not df.empty and '報到狀態' in df.columns:
        df['報到狀態'] = df['報到狀態'].astype(str).str.upper() == 'TRUE'
    return df

if 'attendees' not in st.session_state:
    st.session_state.attendees = load_data()

df = st.session_state.attendees

# ==========================================
# 4. 頂部儀表板：極簡報到統計
# ==========================================
st.title("📱 會議報到系統")

if not df.empty:
    total_people = len(df)
    checked_in_people = df['報到狀態'].sum()
    check_in_rate = (checked_in_people / total_people) * 100 if total_people > 0 else 0
else:
    total_people = checked_in_people = check_in_rate = 0

col1, col2, col3 = st.columns(3)
col1.metric("總人數", f"{total_people}")
col2.metric("已報到", f"{checked_in_people}")
col3.metric("報到率", f"{check_in_rate:.0f}%")

st.divider()

# ==========================================
# 5. 行動端優化介面：三大功能分頁
# ==========================================
tab_checkin, tab_manage, tab_add = st.tabs(["📱 快速報到", "📋 名單管理", "➕ 臨時新增"])

# ------------------------------------------
# 分頁 A：快速報到 (下拉選單 + 代出席變更)
# ------------------------------------------
with tab_checkin:
    if not df.empty:
        # 1. 選擇單位
        units = sorted(df['單位'].astype(str).unique())
        selected_unit = st.selectbox("1️⃣ 請選擇單位", options=["-- 請選擇 --"] + units)
        
        if selected_unit != "-- 請選擇 --":
            # 2. 過濾出該單位的人員，並顯示姓名選單
            unit_df = df[df['單位'] == selected_unit]
            names = unit_df['姓名'].astype(str).tolist()
            selected_name = st.selectbox("2️⃣ 請選擇報到人員", options=["-- 請選擇 --"] + names)
            
            if selected_name != "-- 請選擇 --":
                # 取得選定人員的資料與原始 index
                person_data = unit_df[unit_df['姓名'] == selected_name].iloc[0]
                person_index = unit_df[unit_df['姓名'] == selected_name].index[0]
                
                st.info(f"📍 目前狀態：**{'✅ 已報到' if person_data['報到狀態'] else '🔴 未報到'}**")
                
                if not person_data['報到狀態']:
                    # 3. 變更出席人員選項 (防呆折疊面板)
                    is_substitute = st.checkbox("🔄 換人代為出席 (修改姓名/職稱)")
                    
                    final_name = person_data['姓名']
                    final_title = person_data['職稱']
                    
                    if is_substitute:
                        st.caption("請輸入實際出席人員資訊，系統將自動更新名單：")
                        final_title = st.text_input("實際出席職稱", value=person_data['職稱'], placeholder="例如：專員")
                        final_name = st.text_input("實際出席姓名", value=person_data['姓名'])
                    
                    # 4. 滿版大按鈕確認簽到
                    if st.button("✅ 確認簽到", type="primary", use_container_width=True):
                        current_time = datetime.now().strftime("%H:%M:%S")
                        try:
                            # 雲端寫入 (列數 = index + 2)
                            # 欄位順序：1=單位, 2=職稱, 3=姓名, 4=報到狀態, 5=報到時間
                            if is_substitute:
                                sheet.update_cell(person_index + 2, 2, final_title)
                                sheet.update_cell(person_index + 2, 3, final_name)
                                
                            sheet.update_cell(person_index + 2, 4, "TRUE")
                            sheet.update_cell(person_index + 2, 5, current_time)
                            
                            # 重新載入資料確保完全同步
                            st.session_state.attendees = load_data()
                            st.success(f"✅ {final_name} 報到成功！")
                            st.rerun()
                        except Exception as e:
                            st.error(f"寫入雲端失敗：{e}")
                else:
                    st.button("✅ 此人已完成報到", disabled=True, use_container_width=True)

# ------------------------------------------
# 分頁 B：名單管理 (含搜尋與取消報到)
# ------------------------------------------
with tab_manage:
    search_term = st.text_input("🔍 搜尋姓名或單位 (查詢或取消報到用)")
    show_df = df
    
    if search_term:
        show_df = show_df[show_df['姓名'].astype(str).str.contains(search_term) | 
                          show_df['單位'].astype(str).str.contains(search_term)]
                          
    for index, row in show_df.iterrows():
        # 手機排版：資訊在上，按鈕在下，視覺更清晰
        st.markdown(f"**{row['姓名']}** |  {row['單位']} {row['職稱']}")
        
        if row['報到狀態']:
            col_time, col_btn = st.columns([6, 4], vertical_alignment="center")
            with col_time:
                st.caption(f"🕒 {row['報到時間']}")
            with col_btn:
                if st.button("❌ 取消", key=f"cancel_{index}", use_container_width=True):
                    try:
                        sheet.update_cell(index + 2, 4, "FALSE")
                        sheet.update_cell(index + 2, 5, "")
                        st.session_state.attendees = load_data()
                        st.rerun()
                    except Exception as e:
                        st.error(f"取消失敗：{e}")
        else:
            st.caption("🔴 未報到")
            
        st.divider()

# ------------------------------------------
# 分頁 C：臨時新增人員
# ------------------------------------------
with tab_add:
    st.markdown("### ✍️ 現場報名登錄")
    with st.form("add_new_attendee", clear_on_submit=True):
        st.caption("填寫完畢送出後，系統將自動為該人員完成簽到，並新增於清單最下方。")
        
        new_dept = st.text_input("單位名稱*", placeholder="例如：外部單位")
        col_new1, col_new2 = st.columns(2)
        with col_new1:
            new_title = st.text_input("職稱", placeholder="例如：經理")
        with col_new2:
            new_name = st.text_input("人員姓名*", placeholder="例如：王小明")
            
        submit_new = st.form_submit_button("送出並自動簽到", type="primary", use_container_width=True)
        
        if submit_new:
            if new_dept.strip() and new_name.strip():
                current_time = datetime.now().strftime("%H:%M:%S")
                try:
                    # 雲端寫入欄位順序：單位, 職稱, 姓名, 狀態, 時間
                    sheet.append_row([new_dept, new_title, new_name, "TRUE", current_time])
                    st.session_state.attendees = load_data()
                    st.success(f"✅ 已成功新增【{new_name}】並完成報到！")
                    st.rerun()
                except Exception as e:
                    st.error(f"新增失敗：{e}")
            else:
                st.warning("⚠️ 單位與姓名為必填欄位！")