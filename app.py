import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.title("30分値 → 雛形フォーマット変換アプリ")

# 出力ファイル名入力（必須）
output_filename = st.text_input("出力ファイル名（拡張子 .xlsx は自動で付きます）", value="", help="例: catsapporo_202406")
if not output_filename:
    st.warning("出力ファイル名を入力してください。")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

template_file = "雛形_伊藤忠.xlsx"

def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

def read_uploaded(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, skiprows=5, header=None)
    else:
        xlsx = pd.ExcelFile(file)
        all_sheets = []
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, skiprows=5, header=None)
            df['Sheet'] = sheet_name
            all_sheets.append(df)
        df = pd.concat(all_sheets, ignore_index=True)
    return df

# メイン処理
if uploaded_files and output_filename:
    monthly_data = init_monthly_data()

    for file in uploaded_files:
        df = read_uploaded(file)

        for _, row in df.iterrows():
            try:
                date = pd.to_datetime(row[0], errors='coerce')
                if pd.isnull(date):
                    continue

                mm = date.month
                month_index = mm if mm >= 4 else mm + 12
                key = 'holiday' if is_holiday(date) else 'weekday'

                for
