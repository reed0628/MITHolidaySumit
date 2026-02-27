import streamlit as st
import openpyxl
import random
import io
from datetime import datetime
from openpyxl.cell.cell import MergedCell

# --- 同事名單 ---
EMPLOYEE_LIST = [
    "陳育正 / Reed Chen", "蕭芮淇 / Charlotte Hsiao", "江亞璇 / Joyce Chiang",
    "陳幼慧 / Emily Chen", "高筑音 / Apple Kao", "林耕宇 / Benjamin", "林見松 / Jason Lin"
]

def get_random_time(sh, sm, eh, em):
    total_s = sh * 60 + sm
    total_e = eh * 60 + em
    rnd = random.randint(total_s, total_e)
    return f"{rnd // 60:02d}:{rnd % 60:02d}"

# --- 萬能安全寫入函數 (解決合併儲存格問題) ---
def safe_write(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

def process_excel(file, selected_name, leave_data):
    wb = openpyxl.load_workbook(file)
    # 優先抓取分頁
    ws = wb["海瀧簽到表"] if "海瀧簽到表" in wb.sheetnames else wb.worksheets[0]
    
    # 1. 寫入姓名
    safe_write(ws, 2, 2, f"姓名：  {selected_name}")
    
    # 2. 處理 1號到 31號 (Row 4 ~ 34)
    for row in range(4, 35):
        desc_cell = ws.cell(row=row, column=4)
        desc_val = str(desc_cell.value).strip() if desc_cell.value else ""
        
        date_cell = ws.cell(row=row, column=2)
        if not date_cell.value: continue
        
        try:
            if isinstance(date_cell.value, datetime):
                date_str = date_cell.value.strftime("%m/%d")
            else:
                date_str = str(date_cell.value)[5:10].replace("-", "/")
        except:
            date_str = ""

        # A. 假日畫斜線
        if "假日" in desc_val:
            for col in range(5, 10):
                safe_write(ws, row, col, "/")
            continue

        # B. 工作日填時間
        if "工作" in desc_val:
            on_t = get_random_time(8, 50, 9, 5)
            off_t = get_random_time(18, 0, 18, 10)
            remark = ""

            if date_str in leave_data:
                l = leave_data[date_str]
                remark = f"{l['type']} {l['start']}-{l['end']}"
                # 請假邏輯判斷
                if l['end'] == "12:00":
                    on_t = "13:30"
                elif l['start'] >= "13:30":
                    off_t = l['start']
                
                # 全天請假判斷 (這是剛才斷掉的地方)
                if l['start'] <= "09:00" and l['end'] >= "18:00":
                    on_t, off_t = "請假", "請假"

            safe_write(ws, row, 5, on_t)
            safe_write(ws, row, 7, off_t)
            safe_write(ws, row, 9, remark)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# --- 網頁介面 ---
st.set_page_config(page_title="海瀧出勤工具", layout="centered")
st.title("🚢 海瀧出勤紀錄自動填表")
name_choice = st.selectbox("1. 選擇姓名", EMPLOYEE_LIST)
uploaded_file = st.file_uploader("2. 上傳 Excel 範本", type=["xlsx"])

if uploaded_file:
    if 'leaves' not in st.session_state: st.session_state.leaves = {}
    st.subheader("3. 休假設定")
    c1, c2, c3, c4 = st.columns(4)
    with c1: d_in = st.text_input("日期(MM/DD)", placeholder="02/09")
    with c2: t_in = st.selectbox("假別", ["特休", "事假", "病假", "公假"])
    with c3: s_in = st.text_input("開始", "09:00")
    with c4: e_in = st.text_input("結束", "12:00")
    
    if st.button("➕ 新增"):
        if d_in:
            st.session_state.leaves[d_in] = {"type": t_in, "start": s_in, "end": e_in}
            st.rerun()

    if st.session_state.leaves:
        st.write("目前設定：", st.session_state.leaves)
        if st.button("🗑️ 清空"):
            st.session_state.leaves = {}
            st.rerun()

    if st.button("🚀 生成下載"):
        try:
            final_xlsx = process_excel(uploaded_file, name_choice, st.session_state.leaves)
            st.download_button("💾 點我下載成品", final_xlsx, f"{name_choice.split(' / ')[0]}_出勤表.xlsx")
        except Exception as e:
            st.error(f"錯誤：{e}")
