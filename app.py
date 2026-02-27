import streamlit as st
import pandas as pd
import openpyxl
import random
import io
from datetime import datetime

# --- 姓名名單來源：根據你提供的員工清單 CSV ---
EMPLOYEE_LIST = [
    "陳育正 / Reed Chen",
    "蕭芮淇 / Charlotte Hsiao",
    "江亞璇 / Joyce Chiang",
    "陳幼慧 / Emily Chen",
    "高筑音 / Apple Kao",
    "林耕宇 / Benjamin",
    "林見松 / Jason Lin"
]

def get_random_time(start_h, start_m, end_h, end_m):
    total_start = start_h * 60 + start_m
    total_end = end_h * 60 + end_m
    random_minutes = random.randint(total_start, total_end)
    return f"{random_minutes // 60:02d}:{random_minutes % 60:02d}"

def process_excel(file, selected_name, leave_data):
    wb = openpyxl.load_workbook(file)
    
    # 【核心修正】指定分頁名稱，避免抓錯頁
    try:
        ws = wb["海瀧簽到表"]
    except KeyError:
        # 如果萬一分頁名稱不對，就抓第一張分頁
        ws = wb.worksheets[0]
        st.warning(f"找不到名為『海瀧簽到表』的分頁，程式改為處理：{ws.title}")
    
    # 1. 在 B2 填入選定的姓名
    ws["B2"] = f"姓名：  {selected_name}"
    
    # 2. 處理出勤明細 (從第 4 列到第 34 列)
    for row in range(4, 35):
        desc_cell = ws.cell(row=row, column=4) # D 欄：說明
        desc_val = str(desc_cell.value).strip() if desc_cell.value else ""
        
        date_cell = ws.cell(row=row, column=2) # B 欄：日期
        if not date_cell.value:
            continue
            
        try:
            if isinstance(date_cell.value, datetime):
                date_str = date_cell.value.strftime("%m/%d")
            else:
                date_str = str(date_cell.value)[5:10].replace("-", "/")
        except:
            date_str = ""

        # --- 邏輯 A：國定假日或周末假日 畫斜線 ---
        if "假日" in desc_val:
            for col in range(5, 10): # E, F, G, H, I 欄全部填斜線
                ws.cell(row=row, column=col).value = "/"
            continue

        # --- 邏輯 B：工作日 跑隨機時間 ---
        if "工作日" in desc_val:
            on_time = get_random_time(8, 50, 9, 5)
            off_time = get_random_time(18, 0, 18, 10)
            remark = ""

            if date_str in leave_data:
                leave = leave_data[date_str]
                remark = f"{leave['type']} {leave['start']}-{leave['end']}"
                if leave['end'] == "12:00":
                    on_time = "13:30"
                elif leave['start'] >= "13:30":
                    off_time = leave['start']
                if leave['start'] <= "09:00" and leave['end'] >= "18:00":
                    on_time, off_time = "請假", "請假"

            ws.cell(row=row, column=5).value = on_time # E 上班
            ws.cell(row=row, column=7).value = off_time # G 下班
            ws.cell(row=row, column=9).value = remark   # I 備註

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- Streamlit UI介面 ---
st.set_page_config(page_title="海瀧出勤工具", layout="centered")
st.title("🚢 海瀧出勤紀錄自動填表")

name_choice = st.selectbox("1. 請選擇填表人姓名", EMPLOYEE_LIST)

uploaded_file = st.file_uploader("2. 上傳空白 Excel 範本", type=["xlsx"])

if uploaded_file:
    if 'leaves' not in st.session_state: st.session_state.leaves = {}
    st.subheader("3. 設定休假日期 (非必填)")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1: d_in = st.text_input("日期 (MM/DD)", placeholder="02/09")
    with col2: t_in = st.selectbox("假別", ["特休", "事假", "病假", "公假"])
    with col3: s_in = st.text_input("開始", "09:00")
    with col4: e_in = st.text_input("結束", "12:00")
    
    if st.button("➕ 新增休假"):
        if d_in:
            st.session_state.leaves[d_in] = {"type": t_in, "start": s_in, "end": e_in}
            st.rerun()

    if st.session_state.leaves:
        st.write("已設定休假：", st.session_state.leaves)
        if st.button("🗑️ 清空休假設定"):
            st.session_state.leaves = {}
            st.rerun()

    if st.button("🚀 生成並下載 Excel"):
        final_xlsx = process_excel(uploaded_file, name_choice, st.session_state.leaves)
        st.download_button(
            label="💾 點我下載成品",
            data=final_xlsx,
            file_name=f"{name_choice.split(' / ')[0]}_出勤紀錄表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
