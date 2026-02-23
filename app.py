import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import json

# ==========================================
# 1. 頁面與基本設定
# ==========================================
st.set_page_config(page_title="會議報到工作站", page_icon="📱", layout="centered")

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
    
    if '職稱' not in df.columns:
        df['職稱'] = ""
    if not df.empty and '報到狀態' in df.columns:
        df['報到狀態'] = df['報到狀態'].astype(str).str.upper() == 'TRUE'
    return df

if 'attendees' not in st.session_state:
    st.session_state.attendees = load_data()

df = st.session_state.attendees

# ==========================================
# 4. 頂部儀表板
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

# 初始化搜尋欄的 Session State
if 'search_term' not in st.session_state:
    st.session_state.search_term = ""

# ==========================================
# 5. 三大功能分頁
# ==========================================
tab_checkin, tab_manage, tab_add = st.tabs(["📱 快速報到", "📋 名單管理", "➕ 臨時新增"])

# ------------------------------------------
# 分頁 A：快速報到 
# ------------------------------------------
with tab_checkin:
    if not df.empty:
        def clear_search():
            st.session_state.search_term = ""

        col_search, col_clear = st.columns([8, 2], vertical_alignment="bottom")
        
        with col_search:
            search_mode = st.text_input("🔍 快速搜尋 (輸入姓名或單位)", key="search_term", placeholder="輸入關鍵字...")
        with col_clear:
            st.button("✖ 清除", on_click=clear_search, use_container_width=True)
                
        st.caption("或使用下方選單挑選：")
        
        if search_mode:
            search_df = df[df['姓名'].astype(str).str.contains(search_mode) | 
                           df['單位'].astype(str).str.contains(search_mode)]
            
            if search_df.empty:
                st.warning("找不到符合的人員。")
            else:
                for index, row in search_df.iterrows():
                    st.markdown(f"**{row['姓名']}** | {row['單位']} {row['職稱']}")
                    if row['報到狀態']:
                        st.button("✅ 已完成報到", disabled=True, key=f"s_done_{index}", use_container_width=True)
                    else:
                        if st.button("👉 確認簽到", type="primary", key=f"s_checkin_{index}", use_container_width=True):
                            current_time = datetime.now().strftime("%H:%M:%S")
                            try:
                                sheet.update_cell(index + 2, 4, "TRUE")
                                sheet.update_cell(index + 2, 5, current_time)
                                st.session_state.attendees = load_data()
                                st.toast(f"✅ {row['姓名']} 報到成功！", icon="✅")
                                st.rerun()
                            except Exception as e:
                                st.error(f"寫入雲端失敗：{e}")
                    st.divider()
        else:
            units = sorted(df['單位'].astype(str).unique())
            selected_unit = st.selectbox("1️⃣ 請選擇單位", options=["-- 請選擇 --"] + units)
            
            if selected_unit != "-- 請選擇 --":
                unit_df = df[df['單位'] == selected_unit]
                names = unit_df['姓名'].astype(str).tolist()
                selected_name = st.selectbox("2️⃣ 請選擇報到人員", options=["-- 請選擇 --"] + names)
                
                if selected_name != "-- 請選擇 --":
                    person_data = unit_df[unit_df['姓名'] == selected_name].iloc[0]
                    person_index = unit_df[unit_df['姓名'] == selected_name].index[0]
                    
                    st.info(f"📍 目前狀態：**{'✅ 已報到' if person_data['報到狀態'] else '🔴 未報到'}**")
                    
                    if not person_data['報到狀態']:
                        is_substitute = st.checkbox("🔄 換人代為出席 (修改姓名/職稱)")
                        
                        final_name = person_data['姓名']
                        final_title = person_data['職稱']
                        
                        if is_substitute:
                            st.caption("請輸入實際出席人員資訊，系統將自動更新名單：")
                            final_title = st.text_input("實際出席職稱", value=person_data['職稱'])
                            final_name = st.text_input("實際出席姓名", value=person_data['姓名'])
                        
                        if st.button("✅ 確認簽到", type="primary", use_container_width=True):
                            current_time = datetime.now().strftime("%H:%M:%S")
                            try:
                                if is_substitute:
                                    sheet.update_cell(person_index + 2, 2, final_title)
                                    sheet.update_cell(person_index + 2, 3, final_name)
                                    
                                sheet.update_cell(person_index + 2, 4, "TRUE")
                                sheet.update_cell(person_index + 2, 5, current_time)
                                
                                st.session_state.attendees = load_data()
                                st.toast(f"✅ {final_name} 報到成功！", icon="✅")
                                st.rerun()
                            except Exception as e:
                                st.error(f"寫入雲端失敗：{e}")
                    else:
                        st.button("✅ 此人已完成報到", disabled=True, use_container_width=True)

# ------------------------------------------
# 分頁 B：名單管理 (純視覺確認打勾 + 獨立撤銷按鈕)
# ------------------------------------------
with tab_manage:
    st.caption("☑️ 右方勾選框僅供**內部核對使用**（預設為空，打勾不影響雲端紀錄）。若需取消報到請點擊「✖ 撤銷」。")
    
    checked_in_df = df[df['報到狀態'] == True]
    
    search_manage = st.text_input("🔍 搜尋已報到名單", key="manage_search")
    if search_manage:
        checked_in_df = checked_in_df[checked_in_df['姓名'].astype(str).str.contains(search_manage) | 
                                      checked_in_df['單位'].astype(str).str.contains(search_manage)]
    
    if checked_in_df.empty:
        st.info("目前無符合條件的已報到紀錄。")
        
    for index, row in checked_in_df.iterrows():
        # 💡 將版面切分為三塊：資訊佔6份、打勾佔2份、撤銷佔2份
        col_info, col_checkbox, col_cancel = st.columns([6, 2, 2], vertical_alignment="center")
        
        with col_info:
            st.markdown(f"**{row['姓名']}** |  {row['單位']} {row['職稱']}")
            st.caption(f"🕒 報到時間：{row['報到時間']}")
            
        with col_checkbox:
            # 💡 純視覺用途的勾選框 (預設為 False，不觸發任何資料庫更新)
            st.checkbox("已確認", value=False, key=f"confirm_cb_{index}")
            
        with col_cancel:
            # 💡 獨立的撤銷按鈕，保留取消報到的功能
            if st.button("✖ 撤銷", key=f"btn_cancel_{index}"):
                try:
                    sheet.update_cell(index + 2, 4, "FALSE")
                    sheet.update_cell(index + 2, 5, "")
                    st.session_state.attendees = load_data()
                    st.toast(f"⚠️ 已撤銷 {row['姓名']} 的報到紀錄")
                    st.rerun()
                except Exception as e:
                    st.error(f"撤銷失敗：{e}")
                    
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
                    sheet.append_row([new_dept, new_title, new_name, "TRUE", current_time])
                    st.session_state.attendees = load_data()
                    st.success(f"✅ 已成功新增【{new_name}】並完成報到！")
                    st.rerun()
                except Exception as e:
                    st.error(f"新增失敗：{e}")
            else:
                st.warning("⚠️ 單位與姓名為必填欄位！")