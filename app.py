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

# --- 新增：萬能安全寫入函數 ---
def safe_write(ws, r, c, value):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        # 如果是合併格，找出該區域的左上角主格
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = value
                return
    else:
        cell.value = value

def process_excel(file, selected_name, leave_data):
    wb = openpyxl.load_workbook(file)
    # 優先抓取分頁，防呆機制
    ws = wb["海瀧簽到表"] if "海瀧簽到表" in wb.sheetnames else wb.worksheets[0]
    
    # 1. 寫入姓名 (B2)
    safe_write(ws, 2, 2, f"姓名：  {selected_name}")
    
    # 2. 處理 1號到 31號 (Row 4 ~ 34)
    for row in range(4, 35):
        desc_cell = ws.cell(row=row, column=4) # D 欄：說明
        desc_val = str(desc_cell.value).strip() if desc_cell.value else ""
        
        # 讀取日期格 (B 欄) 做比對
        date_cell = ws.cell(row=row, column=2)
        if not date_cell.value: continue
        
        try:
            date_str = date_cell.value.strftime("%m/%d") if isinstance(date_cell.value, datetime) else str(date_cell.value)[5:10].replace("-", "/")
        except: date_str = ""

        # A. 只要說明包含「假日」，全部畫斜線
        if "假日" in desc_val:
            for col in range(5, 10): # E(5) 到 I(9)
                safe_write(ws, row, col, "/")
            continue

        # B. 只要說明包含「工作」，自動填時間
        if "工作" in desc_val:
            on_t, off_t, remark = get_random_time(8, 50, 9, 5), get_random_time(18, 0, 18, 10), ""

            if date_str in leave_data:
                l = leave_data[date_str]
                remark = f"{l['type']} {l['start']}-{l['end']}"
                if l['end'] == "12:00": on_t = "13:30"
                elif l['start'] >= "13:30": off_t = l['start']
                if l['start'] <= "09:0
