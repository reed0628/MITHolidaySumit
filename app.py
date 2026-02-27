import streamlit as st
import pandas as pd
import openpyxl
import random
import io
from datetime import datetime

# --- 預設名單 ---
EMPLOYEE_LIST = [
    "陳育正 / Reed Chen",
    "蕭芮淇 / Charlotte Hsiao",
    "江亞璇 / Joyce Chiang",
    "陳幼慧 / Emily Chen",
    "高筑音 / Apple Kao"
]

def get_random_time(start_h, start_m, end_h, end_m):
    total_start = start_h * 60 + start_m
    total_end = end_h * 60 + end_m
    random_minutes = random.randint(total_start, total_end)
    return f"{random_minutes // 60:02d}:{random_minutes % 60:02d}"

def process_excel(file, selected_name, leave_data):
    wb = openpyxl.load_workbook(file)
    
    # 【關鍵修正】精準指定分頁名稱，避免抓錯頁
    try:
        ws = wb["海瀧簽到表"]
    except KeyError:
        # 如果找不到該名稱，就抓第一張表
        ws = wb.worksheets[0]
        st.warning(f"找不到『海瀧簽到表』分頁，程式已自動抓取第一張表：{ws.title}")
    
    # 1. 填入姓名 (在 B2 儲存格)
    ws["B2"] = f"姓名：  {selected_name}"
    
    # 2. 開始處理每一列 (從第 4 列開始)
    for row in range(4, 35):
        # 讀取「說明」欄位 (D 欄，Index 4)
        desc_cell = ws.cell(row=row, column=4)
        desc_val = str(desc_cell.value).strip() if desc_cell.value else ""
        
        # 讀取「日期」欄位 (B 欄)
        date_cell = ws.cell(row=row, column=2)
        if not date_cell.value:
            continue
            
        # 處理日期格式比對
        try:
            if isinstance(date_cell.value, datetime):
                date_str = date_cell.value.strftime("%m/%d")
            else:
                date_str = str(date_cell.value)[5:10].replace("-", "/")
        except:
            date_str = ""

        # --- 邏輯 A：假日/國定假日 畫斜線 ---
        # 只要說明欄位包含 "假日" 二字就畫斜線
        if "假日" in desc_val:
            for col in [5, 6, 7, 8, 9]: # E, F, G, H, I 欄
                ws.cell(row=row, column=col).value = "/"
            continue

        # --- 邏輯 B：工作日 生成時間 ---
        if "工作日" in desc_val:
            on_time = get_random_time(8, 50, 9, 5)
            off_time = get_random_time(18, 0, 18, 10)
            remark = ""

            # 處理休假
            if date_str in leave_data:
                leave = leave_data[date_str]
                remark = f"{leave['type']} {leave['start']}-{leave['end']}"
                
                if leave['end'] == "12:00":
                    on_time = "13:30"
                elif leave['start'] >= "13:30":
                    off_time = leave['start']
                if leave['start'] <= "09:00" and leave['end'] >= "18:00":
                    on_time, off_time = "請假", "請假"

            # 寫入 (E:5, G:7, I:9)
            ws.cell(row=row, column=5).value = on_time
            ws.cell(row=row, column=7).value = off_time
            ws.cell(row=row, column=9).value = remark

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- 網頁介面省略 (保持不變) ---
st.set_page_config(page_title="海瀧出勤工具", layout="centered")
st.title("🚢 出勤紀錄表自動生成器")
selected_name = st.selectbox("1. 請選擇你的姓名", EMPLOYEE_LIST)
uploaded_file = st.file_uploader("2. 上傳空白 Excel 範本", type=["xlsx"])

if uploaded_file:
    if 'leaves' not in st.session_state: st.session_state.leaves = {}
    st.subheader("3. 設定休假日期")
    c1, c2, c3, c4 = st.columns(4)
    with c1: d_in = st.text_input("日期 (MM/DD)", placeholder="02/09")
    with c2: t_in = st.selectbox("假別", ["特休", "事假", "病假", "公假"])
    with c3: s_in = st.text_input("開始", "09:00")
    with c4: e_in = st.text_input("結束", "12:00")
    
    if st.button("➕ 新增休假"):
        if d_in:
            st.session_state.leaves[d_in] = {"type": t_in, "start": s_in, "end": e_in}
            st.rerun()

    if st.session_state.leaves:
        st.write("已設定休假：", st.session_state.leaves)
        if st.button("🗑️ 清除所有休假"):
            st.session_state.leaves = {}
            st.rerun()

    if st.button("🚀 生成並下載"):
        final_file = process_excel(uploaded_file, selected_name, st.session_state.leaves)
        st.download_button(
            label="💾 點我下載成品",
            data=final_file,
            file_name=f"{selected_name}_出勤表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
