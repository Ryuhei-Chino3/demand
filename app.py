import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

template_file = "雛形_伊藤忠.xlsx"

# 祝日・土日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 月別初期化（4〜翌3月：4→4, 3→15）
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

# アップロードファイル読み込み
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

# 実行ブロック
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

            for i in range(1, 49):  # 1〜48列（30分値）
                if i >= len(column_names):
                    continue
                val = pd.to_numeric(row[column_names[i]], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

    # 雛形読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 平日エリア（E列:4月 → P列:翌年3月）
    for m in range(4, 16):
        col_index = 1 + (m - 4) + 4  # E列(5) = 4月 → F, G, H...
        col_letter = get_column_letter(col_index)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['weekday'][m][i]

    # 休日エリア（S列:4月 → AD列:翌年3月）← 修正ポイント
    for m in range(4, 16):
        col_index = 1 + (m - 4) + 18  # S列(19) = 4月 → T, U...
        col_letter = get_column_letter(col_index)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['holiday'][m][i]

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
