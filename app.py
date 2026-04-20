import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import json

# ==========================================
# 1. 頁面設定
# ==========================================
st.set_page_config(page_title="會議報到工作站", page_icon="📋", layout="wide")

# ==========================================
# 2. Google Sheets 連線
# ==========================================
@st.cache_resource
def init_gsheets_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        secret_dict = json.loads(st.secrets["gcp_json"])
        creds = Credentials.from_service_account_info(secret_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_key("1Ezc7IVTQJF76pSCrZEsF2n9vW_dBEPDyYrsx2-jwZOI")
    except Exception as e:
        st.error(f"資料庫連線失敗：{e}")
        st.stop()

def get_sheet(spreadsheet, name):
    try:
        return spreadsheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=name, rows=500, cols=10)
        ws.append_row(["姓名", "單位", "級職", "報到狀態", "報到時間"])
        return ws

spreadsheet = init_gsheets_client()

# ==========================================
# 3. 側邊欄：時段 + 功能導覽
# ==========================================
SESSIONS = {"🌅 上午場": "上午", "🌇 下午場": "下午"}

with st.sidebar:
    st.header("⏰ 時段選擇")
    selected_label = st.radio(
        "", options=list(SESSIONS.keys()), index=0, label_visibility="collapsed"
    )
    selected_session = SESSIONS[selected_label]

    st.divider()

    st.header("📌 功能選擇")
    page = st.radio(
        "",
        options=["✅ 簽到登錄", "📋 簽到管理", "➕ 臨時簽到"],
        index=0,
        label_visibility="collapsed",
    )

sheet = get_sheet(spreadsheet, selected_session)

# ==========================================
# 4. 動態讀取欄位對照（適配新舊 Sheet 結構）
# ==========================================
col_map_key = f"col_map_{selected_session}"
if col_map_key not in st.session_state:
    headers = sheet.row_values(1)
    st.session_state[col_map_key] = {h: i + 1 for i, h in enumerate(headers)}

col_map = st.session_state[col_map_key]

COL_NAME   = col_map.get("姓名", 1)
COL_UNIT   = col_map.get("單位", 2)
COL_RANK   = col_map.get("級職", None)   # 舊版 Sheet 可能沒有此欄
COL_STATUS = col_map.get("報到狀態", 3)
COL_TIME   = col_map.get("報到時間", 4)
HAS_RANK   = COL_RANK is not None

# ==========================================
# 5. 資料載入
# ==========================================
def load_data(ws):
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    if not df.empty and "報到狀態" in df.columns:
        df["報到狀態"] = df["報到狀態"].astype(str).str.upper() == "TRUE"
    if not df.empty and "級職" not in df.columns:
        df["級職"] = ""
    if not df.empty and "報到時間" not in df.columns:
        df["報到時間"] = ""
    return df

def refresh_data():
    """重新從雲端拉取資料，並更新欄位對照"""
    st.session_state[session_key] = load_data(sheet)
    headers = sheet.row_values(1)
    st.session_state[col_map_key] = {h: i + 1 for i, h in enumerate(headers)}

session_key = f"attendees_{selected_session}"
if session_key not in st.session_state:
    st.session_state[session_key] = load_data(sheet)

df_full = st.session_state[session_key]

# ==========================================
# 6. 統計儀表板（頂部，所有頁面共用）
# ==========================================
st.title(f"📋 現場會議報到　｜　{selected_label}")

total   = len(df_full) if not df_full.empty else 0
checked = int(df_full["報到狀態"].sum()) if not df_full.empty else 0
rate    = (checked / total * 100) if total > 0 else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("總人數", f"{total} 人")
c2.metric("已報到", f"{checked} 人")
c3.metric("未報到", f"{total - checked} 人")
c4.metric("報到率", f"{rate:.1f} %")
st.divider()

# ==========================================
# 輔助函式：簽到 / 撤銷
# ==========================================
def do_checkin(index, name):
    now = datetime.now().strftime("%H:%M:%S")
    try:
        sheet.update_cell(index + 2, COL_STATUS, "TRUE")
        sheet.update_cell(index + 2, COL_TIME, now)
        st.session_state[session_key].at[index, "報到狀態"] = True
        st.session_state[session_key].at[index, "報到時間"] = now
        st.toast(f"✅ {name} 報到成功！", icon="✅")
        st.rerun()
    except Exception as e:
        st.error(f"簽到失敗：{e}")

def do_undo(index, name):
    try:
        sheet.update_cell(index + 2, COL_STATUS, "FALSE")
        sheet.update_cell(index + 2, COL_TIME, "")
        st.session_state[session_key].at[index, "報到狀態"] = False
        st.session_state[session_key].at[index, "報到時間"] = ""
        st.toast(f"↩ {name} 簽到已撤銷")
        st.rerun()
    except Exception as e:
        st.error(f"撤銷失敗：{e}")

def row_rank(row):
    r = str(row.get("級職", "")).strip()
    return f" · {r}" if r else ""

# ==========================================
# 頁面 A：簽到登錄
# ==========================================
if page == "✅ 簽到登錄":
    st.subheader("✅ 簽到登錄")

    if df_full.empty:
        st.info("目前名冊為空，請先匯入 Google Sheets 名單，或使用「臨時簽到」新增人員。")
    else:
        # 三欄級聯篩選
        fc1, fc2, fc3 = st.columns(3)

        with fc1:
            unit_list = ["全部單位"] + sorted(df_full["單位"].astype(str).unique().tolist())
            sel_unit = st.selectbox("🏢 單位", unit_list)

        with fc2:
            tmp = df_full if sel_unit == "全部單位" else df_full[df_full["單位"].astype(str) == sel_unit]
            rank_vals = tmp["級職"].astype(str).replace("", "（未填）").unique().tolist()
            rank_list = ["全部級職"] + sorted(rank_vals)
            sel_rank = st.selectbox("🎖 級職", rank_list)

        with fc3:
            tmp2 = tmp.copy()
            if sel_rank != "全部級職":
                rv = "" if sel_rank == "（未填）" else sel_rank
                tmp2 = tmp2[tmp2["級職"].astype(str) == rv]
            name_list = ["全部人員"] + tmp2["姓名"].astype(str).tolist()
            sel_name = st.selectbox("👤 姓名", name_list)

        # 套用篩選
        result = df_full.copy()
        if sel_unit != "全部單位":
            result = result[result["單位"].astype(str) == sel_unit]
        if sel_rank != "全部級職":
            rv = "" if sel_rank == "（未填）" else sel_rank
            result = result[result["級職"].astype(str) == rv]
        if sel_name != "全部人員":
            result = result[result["姓名"].astype(str) == sel_name]

        st.caption(f"顯示 {len(result)} 筆 / 共 {total} 筆")

        for index, row in result.iterrows():
            col_no, col_info, col_edit, col_action = st.columns(
                [1, 5, 3, 3], vertical_alignment="center"
            )

            with col_no:
                st.caption(f"#{index + 1}")

            with col_info:
                st.markdown(f"**{row['姓名']}** | {row['單位']}{row_rank(row)}")
                if row["報到狀態"]:
                    st.caption(f"🕒 {row['報到時間']}")
                else:
                    st.caption("🔴 未報到")

            with col_edit:
                with st.expander("✏️ 姓名變更"):
                    new_name_val = st.text_input(
                        "新姓名",
                        value=str(row["姓名"]),
                        key=f"nn_{index}",
                        label_visibility="collapsed",
                    )
                    if st.button("確認變更", key=f"rn_{index}", use_container_width=True):
                        n = new_name_val.strip()
                        if n and n != str(row["姓名"]):
                            try:
                                sheet.update_cell(index + 2, COL_NAME, n)
                                st.session_state[session_key].at[index, "姓名"] = n
                                st.toast(f"✅ 姓名已更新為「{n}」")
                                st.rerun()
                            except Exception as e:
                                st.error(f"更新失敗：{e}")

            with col_action:
                if not row["報到狀態"]:
                    if st.button(
                        "👉 點此簽到",
                        key=f"ci_{index}",
                        use_container_width=True,
                        type="primary",
                    ):
                        do_checkin(index, row["姓名"])
                else:
                    b1, b2 = st.columns(2)
                    with b1:
                        st.button(
                            "✅ 已完成",
                            key=f"done_{index}",
                            disabled=True,
                            use_container_width=True,
                        )
                    with b2:
                        if st.button("↩ 撤銷", key=f"undo_{index}", use_container_width=True):
                            do_undo(index, row["姓名"])

            st.divider()

# ==========================================
# 頁面 B：簽到管理
# ==========================================
elif page == "📋 簽到管理":
    st.subheader("📋 簽到管理")

    btn_r, btn_s, _ = st.columns([2, 2, 6])
    with btn_r:
        if st.button("🔄 重新整理", use_container_width=True):
            refresh_data()
            st.rerun()

    if df_full.empty:
        st.info("目前沒有資料。")
    else:
        # 全選控制
        select_all_checked = st.checkbox("☑ 全選已報到人員（可批次撤銷）")
        selected_indices = []

        # 欄位標題列
        h0, h1, h2, h3, h4, h5 = st.columns([1, 1, 2, 3, 2, 2])
        h0.caption("勾選")
        h1.caption("序號")
        h2.caption("姓名")
        h3.caption("單位 · 級職")
        h4.caption("狀態")
        h5.caption("操作")
        st.markdown("---")

        # 僅顯示已報到人員，依原始 Excel 順序
        df_checked = df_full[df_full["報到狀態"] == True]
        if df_checked.empty:
            st.info("目前尚無已報到人員。")
        for index, row in df_checked.iterrows():
            c0, c1, c2, c3, c4, c5 = st.columns(
                [1, 1, 2, 3, 2, 2], vertical_alignment="center"
            )

            with c0:
                default_val = select_all_checked and bool(row["報到狀態"])
                is_sel = st.checkbox(
                    "",
                    key=f"sel_{index}",
                    value=default_val,
                    label_visibility="collapsed",
                )
                if is_sel:
                    selected_indices.append(index)

            with c1:
                st.caption(f"#{index + 1}")

            with c2:
                st.markdown(f"**{row['姓名']}**")

            with c3:
                st.caption(f"{row['單位']}{row_rank(row)}")

            with c4:
                if row["報到狀態"]:
                    st.markdown("✅ 已報到")
                    st.caption(str(row["報到時間"]))
                else:
                    st.markdown("🔴 未報到")

            with c5:
                if st.button(
                    "↩ 撤銷", key=f"mgmt_undo_{index}", use_container_width=True
                ):
                    do_undo(index, row["姓名"])

        # 批次撤銷
        if selected_indices:
            st.divider()
            if st.button(
                f"↩ 批次撤銷選取的 {len(selected_indices)} 筆",
                type="secondary",
                use_container_width=True,
            ):
                failed = []
                for idx in selected_indices:
                    try:
                        sheet.update_cell(idx + 2, COL_STATUS, "FALSE")
                        sheet.update_cell(idx + 2, COL_TIME, "")
                        st.session_state[session_key].at[idx, "報到狀態"] = False
                        st.session_state[session_key].at[idx, "報到時間"] = ""
                    except Exception as e:
                        failed.append(f"#{idx+1}: {e}")
                if failed:
                    st.error("部分撤銷失敗：\n" + "\n".join(failed))
                else:
                    st.toast(f"↩ 已批次撤銷 {len(selected_indices)} 筆")
                st.rerun()

# ==========================================
# 頁面 C：臨時簽到
# ==========================================
elif page == "➕ 臨時簽到":
    st.subheader("➕ 臨時簽到")
    st.caption("新增不在原始名冊的臨時人員，送出後自動完成簽到並寫入 Google Sheets。")

    with st.form("walk_in_form", clear_on_submit=True):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            wi_unit = st.text_input("🏢 單位 ＊", placeholder="請輸入單位名稱")
        with fc2:
            wi_rank = st.text_input("🎖 職稱", placeholder="可留空")
        with fc3:
            wi_name = st.text_input("👤 姓名 ＊", placeholder="請輸入姓名")

        submitted = st.form_submit_button(
            "➕ 新增並完成簽到", type="primary", use_container_width=True
        )

    if submitted:
        if not wi_unit.strip() or not wi_name.strip():
            st.warning("⚠️ 單位和姓名為必填欄位，請填寫後再送出。")
        else:
            now = datetime.now().strftime("%H:%M:%S")
            try:
                if HAS_RANK:
                    row_data = [wi_name.strip(), wi_unit.strip(), wi_rank.strip(), "TRUE", now]
                else:
                    row_data = [wi_name.strip(), wi_unit.strip(), "TRUE", now]
                sheet.append_row(row_data)

                new_row_df = pd.DataFrame([{
                    "姓名": wi_name.strip(),
                    "單位": wi_unit.strip(),
                    "級職": wi_rank.strip(),
                    "報到狀態": True,
                    "報到時間": now,
                }])
                st.session_state[session_key] = pd.concat(
                    [st.session_state[session_key], new_row_df], ignore_index=True
                )
                st.success(
                    f"✅ **{wi_name.strip()}**（{wi_unit.strip()}）已新增並完成簽到！"
                )
                st.rerun()
            except Exception as e:
                st.error(f"新增失敗：{e}")

    # 顯示本時段已報到名單（最下方供確認）
    st.divider()
    st.caption("📋 本時段完整名單（含臨時人員，按原始順序）")
    if not df_full.empty:
        preview = df_full[["姓名", "單位", "級職", "報到狀態", "報到時間"]].copy()
        preview.index = range(1, len(preview) + 1)
        preview["報到狀態"] = preview["報到狀態"].map({True: "✅ 已報到", False: "🔴 未報到"})
        st.dataframe(preview, use_container_width=True)
    else:
        st.info("目前名單為空。")
