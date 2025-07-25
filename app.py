import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import deepcopy

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

template_file = "雛形_伊藤忠.xlsx"

# 曜日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 空の月別データ構造を作成（4月〜翌年3月 → 4〜15で扱う）
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

# ファイル読込関数
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
            month_index = mm if mm >= 4 else mm + 12  # 1〜3月は+12して13〜15
            key = 'holiday' if is_holiday(date) else 'weekday'

            for i in range(1, 49):  # 1列目から48列目（0は日付）
                if i >= len(column_names):
                    continue
                colname = column_names[i]
                val = pd.to_numeric(row[colname], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

    # 雛形ファイルを読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 平日データ書き込み（C列＝4月、D列＝5月、E列＝6月 ...）
    for m in range(4, 16):
        col_index = (m - 4) + 3  # C列から開始 = 3列目
        col = get_column_letter(col_index)
        for i in range(48):
            ws[f"{col}{4+i}"] = monthly_data['weekday'][m][i]

    # 休日データ書き込み（Q列＝4月、R列＝5月 ...）
    for m in range(4, 16):
        col_index = (m - 4) + 17  # Q列から開始 = 17列目
        col = get_column_letter(col_index)
        for i in range(48):
            ws[f"{col}{4+i}"] = monthly_data['holiday'][m][i]

    # 出力用に保存
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name="output_koma_format.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
