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
    total_s, total_e = sh * 60 + sm, eh * 60 + em
    rnd = random.randint(total_s, total_e)
    return f"{rnd // 60:02d}:{rnd % 60:02d}"

def safe_write(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

def process_excel(file):
    # 【關鍵破解】讀取兩次檔案
    # wb_read：用 data_only=True 讀取，這樣才能看到公式計算出來的「工作日」三個字
    wb_read = openpyxl.load_workbook(file, data_only=True)
    # wb_write：正常讀取，用來填寫時間並存檔，確保不破壞原本的公式跟格式
    wb_write = openpyxl.load_workbook(file)
    
    # 抓取分頁
    sheet_name = "海瀧簽到表" if "海瀧簽到表" in wb_write.sheetnames else wb_write.sheetnames[0]
    ws_read = wb_read[sheet_name]
    ws_write = wb_write[sheet_name]
    
    # 1. 寫入姓名 (B3)
    safe_write(ws_write, 3, 2, f"姓名：  {st.session_state.selected_name}")
    
    # 2. 自動尋找資料起始列 (找「序號」)
    start_row = 5
    for r in range(1, 10):
        if "序號" in str(ws_read.cell(row=r, column=1).value):
            start_row = r + 1
            break

    # 3. 處理出勤明細
    for row in range(start_row, start_row + 31):
        # 【重點】從 ws_read (唯讀版) 抓取資料，才能避開公式
        desc_cell = ws_read.cell(row=row, column=4) # D 欄
        if desc_cell.value is None: continue
        
        desc_val = str(desc_cell.value).strip()
        date_cell = ws_read.cell(row=row, column=2) # B 欄
        
        try:
            d_val = date_cell.value
            if isinstance(d_val, datetime):
                date_str = d_val.strftime("%m/%d")
            elif "/" in str(d_val):
                date_str = str(d_val).strip()
            else:
                date_str = str(d_val)[5:10].replace("-", "/")
        except:
            date_str = ""

        # A. 假日畫斜線 -> 寫入到 ws_write
        if "假日" in desc_val:
            for col in range(5, 10):
                safe_write(ws_write, row, col, "/")
            continue

        # B. 工作日填時間 -> 寫入到 ws_write
        if "工作" in desc_val:
            on_t = get_random_time(8, 50, 9, 5)
            off_t = get_random_time(18, 0, 18, 10)
            remark = ""

            if date_str in st.session_state.leaves:
                l = st.session_state.leaves[date_str]
                remark = f"{l['type']} {l['start']}-{l['end']}"
                if l['end'] == "12:00":
                    on_t = "13:30"
                elif l['start'] >= "13:30":
                    off_t = l['start']
                if l['start'] <= "09:00" and l['end'] >= "18:00":
                    on_t, off_t = "請假", "請假"

            safe_write(ws_write, row, 5, on_t)
            safe_write(ws_write, row, 7, off_t)
            safe_write(ws_write, row, 9, remark)

    output = io.BytesIO()
    wb_write.save(output)
    return output.getvalue()

# --- 網頁介面 ---
st.set_page_config(page_title="海瀧出勤工具", layout="centered")
st.title("🚢 海瀧出勤紀錄自動填表")

# 把姓名存進 session_state 以便全域讀取
st.session_state.selected_name = st.selectbox("1. 選擇姓名", EMPLOYEE_LIST)

uploaded_file = st.file_uploader("2. 上傳 Excel 範本", type=["xlsx"])

if uploaded_file:
    if 'leaves' not in st.session_state: st.session_state.leaves = {}
    st.subheader("3. 休假設定")
    c1, c2, c3, c4 = st.columns(4)
    with c1: d_in = st.text_input("日期(MM/DD)", placeholder="02/09")
    with c2: t_in = st.selectbox("假別", ["特休", "事假", "病假", "公假"])
    with c3: s_in = st.text_input("開始", "09:00")
    with c4: e_in = st.text_input("結束", "12:00")
    
    if st.button("➕ 新增休假"):
        if d_in:
            st.session_state.leaves[d_in] = {"type": t_in, "start": s_in, "end": e_in}
            st.rerun()

    if st.session_state.leaves:
        st.write("已設定休假：", st.session_state.leaves)
        if st.button("🗑️ 清空所有設定"):
            st.session_state.leaves = {}
            st.rerun()

    if st.button("🚀 生成並下載"):
        try:
            # 現在只要傳 file 就好，因為姓名和假單已經透過 session_state 讀取
            final_xlsx = process_excel(uploaded_file)
            download_name = st.session_state.selected_name.split(' / ')[0]
            st.download_button("💾 下載成果 Excel", final_xlsx, f"{download_name}_出勤表.xlsx")
        except Exception as e:
            st.error(f"錯誤：{e}")
