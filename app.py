import streamlit as st
import pandas as pd
import openpyxl
import random
from datetime import datetime, time
import io

# --- 核心邏輯函數 ---
def get_random_time(start_h, start_m, end_h, end_m):
    # 生成範圍內的隨機時間字串
    total_start = start_h * 60 + start_m
    total_end = end_h * 60 + end_m
    random_minutes = random.randint(total_start, total_end)
    return f"{random_minutes // 60:02d}:{random_minutes % 60:02d}"

def process_excel(file, leave_data):
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    
    # 從第 4 列開始處理
    for row in range(4, 35):
        date_val = ws.cell(row=row, column=2).value  # B 欄：日期
        desc_val = ws.cell(row=row, column=4).value  # D 欄：說明
        
        if not date_val or desc_val != "工作日":
            continue
            
        # 取得日期字串 (例如 02/09)
        date_str = str(date_val)[5:10].replace("-", "/")
        
        on_time = get_random_time(8, 50, 9, 5) # 預設 08:50-09:05
        off_time = get_random_time(18, 0, 18, 10) # 預設 18:00-18:10
        remark = ""

        # 處理休假邏輯
        if date_str in leave_data:
            leave = leave_data[date_str]
            remark = f"{leave['type']} {leave['start']}-{leave['end']}"
            
            # 判斷休假對時間的影響
            if leave['start'] == "09:00" and leave['end'] == "12:00":
                on_time = "13:30"
            elif leave['start'] >= "13:30":
                off_time = leave['start']
            elif leave['start'] <= "09:00" and leave['end'] >= "18:00":
                on_time, off_time = "請假", "請假"

        # 寫入 Excel (E:上班, G:下班, I:備註)
        ws.cell(row=row, column=5).value = on_time
        ws.cell(row=row, column=7).value = off_time
        ws.cell(row=row, column=9).value = remark

    # 將結果存入記憶體並回傳
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- Streamlit 網頁界面 ---
st.title("🚢 出勤紀錄表自動生成器")
st.write("上傳空白表，選好休假，一鍵生成！")

uploaded_file = st.file_uploader("1. 上傳空白 Excel 範本", type=["xlsx"])

if uploaded_file:
    st.success("檔案上傳成功！")
    
    # 讀取日期範圍 (簡單模擬)
    st.subheader("2. 設定休假日期")
    st.info("若當天無休假，請直接跳過。")
    
    # 讓使用者動態新增休假
    if 'leaves' not in st.session_state:
        st.session_state.leaves = {}

    col1, col2, col3, col4 = st.columns(4)
    with col1: date_input = st.text_input("日期 (MM/DD)", placeholder="02/09")
    with col2: type_input = st.selectbox("假別", ["特休", "事假", "病假", "公假"])
    with col3: start_input = st.text_input("開始", value="09:00")
    with col4: end_input = st.text_input("結束", value="12:00")
    
    if st.button("新增此筆休假"):
        st.session_state.leaves[date_input] = {"type": type_input, "start": start_input, "end": end_input}
        st.rerun()

    if st.session_state.leaves:
        st.write("目前的休假清單：", st.session_state.leaves)
        if st.button("清除所有休假"):
            st.session_state.leaves = {}
            st.rerun()

    # 3. 生成按鈕
    if st.button("3. 生成並下載出勤表"):
        result = process_excel(uploaded_file, st.session_state.leaves)
        st.download_button(
            label="點我下載成品",
            data=result,
            file_name=f"已填寫出勤表_{datetime.now().strftime('%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )