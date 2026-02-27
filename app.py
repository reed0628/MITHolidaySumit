import streamlit as st
import pandas as pd
import openpyxl
import random
import io
from datetime import datetime

# --- 姓名名單 ---
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
    # 讀取檔案，明確設定 data_only=False 以保留公式（如果有）
    wb = openpyxl.load_workbook(file)
    
    # 優先抓取「海瀧簽到表」，抓不到就抓第一張
    try:
        ws = wb["海瀧簽到表"]
    except KeyError:
        ws = wb.worksheets[0]
    
    # 【關鍵修正】改用 .cell() 寫法，避開 B2 的 AttributeError
    # row=2, column=2 等於 B2
    try:
        ws.cell(row=2, column=2).value = f"姓名：  {selected_name}"
    except Exception as e:
        st.error(f"寫入姓名時發生錯誤：{e}")

    # 處理出勤明細 (Row 4 到 34)
    for row in range(4, 35):
        desc_cell = ws.cell(row=row, column=4) # D 欄
        desc_val = str(desc_cell.value).strip() if desc_cell.value else ""
        
        # 讀取日期 B 欄
        date_cell = ws.cell(row=row, column=2)
        if not date_cell.value:
            continue
            
        try:
            if isinstance(date_cell.value, datetime):
                date_str = date_cell.value.strftime("%m/%d")
            else:
                date_str = str(date_cell.value)[5:10].replace("-", "/")
        except:
            date_str = ""

        # --- 邏輯 A：只要是假日，E, F, G, H, I 全部畫斜線 ---
        if "假日" in desc_val:
            for col in range(5, 10): # E=5, F=6, G=7, H=8, I=9
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

# --- 下方介面保持不變 ---
st.set_page_config(page_title="海瀧出勤工具", layout="centered")
st.title("🚢 海瀧出勤紀錄自動填表")
name_choice = st.selectbox("1. 請選擇填表人姓名", EMPLOYEE_LIST)
uploaded_file = st.file_uploader("2. 上傳空白 Excel 範本", type=["xlsx"])

if uploaded_file:
    if 'leaves' not in st.session_state: st.session_state.leaves = {}
    st.subheader("3. 設定休假日期 (非必填)")
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
        if st.button("🗑️ 清空休假"):
            st.session_state.leaves = {}
            st.rerun()

    if st.button("🚀 生成並下載 Excel"):
        try:
            final_xlsx = process_excel(uploaded_file, name_choice, st.session_state.leaves)
            st.download_button(
                label="💾 點我下載成品",
                data=final_xlsx,
                file_name=f"{name_choice.split(' / ')[0]}_出勤表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as global_e:
            st.error(f"發生程式錯誤：{global_e}")
