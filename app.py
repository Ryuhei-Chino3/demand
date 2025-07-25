import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

template_file = "雛形_伊藤忠.xlsx"

# 曜日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 空の月別データ構造を作成（4月〜翌年3月）
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

# Excel列番号をA, B, ..., Z, AA, AB, ... に変換
def colnum_to_excel_col(n):
    result = ''
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

# ファイル読み込み
def read_uploaded(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=5)
    else:
        xlsx = pd.ExcelFile(file)
        all_sheets = []
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, header=5)
            df['Sheet'] = sheet_name
            all_sheets.append(df)
        df = pd.concat(all_sheets, ignore_index=True)
    return df

# メイン処理
if uploaded_files:
    monthly_data = init_monthly_data()

    for file in uploaded_files:
        df = read_uploaded(file)
        column_names = df.columns.tolist()

        for _, row in df.iterrows():
            date = pd.to_datetime(row[column_names[0]], errors='coerce')
            if pd.isnull(date):
                continue

            mm = date.month
            month_index = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'

            for i in range(1, 49):  # 1列目から48列目（0は日付）
                if i >= len(column_names):
                    continue
                colname = column_names[i]
                val = pd.to_numeric(row[colname], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

    # 雛形読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 平日データ書き込み（C列=3列目からN=14列目）
    for m in range(4, 16):
        col_index = m - 1 + 2  # 4月=3列目
        col_letter = colnum_to_excel_col(col_index)
        for i in range(48):
            ws[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]

    # 休日データ書き込み（Q列=17列目から）
    for m in range(4, 16):
        col_index = 17 + (m - 4)
        col_letter = colnum_to_excel_col(col_index)
        for i in range(48):
            ws[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]

    # 出力
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name="output_koma_format.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
