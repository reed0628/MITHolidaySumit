import streamlit as st
import pandas as pd
import openpyxl
import random
import io
from datetime import datetime
from openpyxl.cell.cell import MergedCell # 導入合併儲存格判斷

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
    wb = openpyxl.load_workbook(file)
    try:
        ws = wb["海瀧簽到表"]
    except KeyError:
        ws = wb.worksheets[0]
    
    # --- 【智慧寫入姓名邏輯】解決 MergedCell 唯讀問題 ---
    target_cell = ws.cell(row=2, column=2) # B2
    name_text = f"姓名：  {selected_name}"

    if isinstance(target_cell, MergedCell):
        # 如果 B2 是合併儲存格的一部分，我們尋找該合併區域的左上角主儲存格
        for merged_range in ws.merged_cells.ranges:
            if target_cell.coordinate in merged_range:
                # 取得合併區域的左上角座標並寫入
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = name_text
                break
    else:
        target_cell.value = name_text

    # 處理出勤明細 (Row 4 到 34)
    for row in range(4, 35):
        desc_cell = ws.cell(row=row, column=4) # D 欄
        desc_val = str(desc_cell.value).strip() if desc_cell.value else ""
        
        date_cell = ws.cell(row=row, column=2) # B 欄
        if not date_cell.value:
            continue
            
        try:
            if isinstance(date_cell.value, datetime):
                date_str = date_cell.value.strftime("%m/%d")
            else:
                date_str = str(date_cell.value)[5:10].replace("-", "/")
        except:
            date_str = ""

        # --- 邏輯 A：假日畫斜線 ---
        if "假日" in desc_val:
            for col in range(5, 10):
                ws.cell(row=row, column=col).value = "/"
            continue

        # --- 邏輯 B：工作日生成時間 ---
        if "工作日" in desc_val:
            on_time = get_random_time(8, 50, 9, 5)
            off_time = get_random
