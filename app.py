import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

template_file = "雛形_伊藤忠.xlsx"

# 土日・祝日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 月別データ初期化（4〜翌3月 → 4〜15）
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

# 入力ファイル読み込み
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
            month_index = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'

            for i in range(1, 49):  # 30分値は1列目〜48列目（A列は日付）
                if i >= len(column_names):
                    continue
                val = pd.to_numeric(row[column_names[i]], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

    # 雛形テンプレート読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 🔵 平日：6月→E列（5列目）, 7月→F列（6列目）
    for m in range(6, 8):  # 対象：6月と7月
        col_index = 4 + (m - 6)  # 6月→4+0=4→E列, 7月→4+1=5→F列
        col_letter = get_column_letter(col_index + 1)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['weekday'][m][i]

    # 🔴 休日：6月→S列（19列目）, 7月→T列（20列目）
    for m in range(6, 8):  # 対象：6月と7月
        col_index = 18 + (m - 6)  # 6月→18+0=18→S列, 7月→19→T列
        col_letter = get_column_letter(col_index + 1)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['holiday'][m][i]

    # ダウンロード処理
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name="output_koma_format.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
