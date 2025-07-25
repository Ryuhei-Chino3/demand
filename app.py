import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook
from copy import deepcopy

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

        # 列名の表示（デバッグ用）
        st.write("読み込んだ列名:", df.columns.tolist())

        # 日付列の自動判定
        date_col = None
        for col in df.columns:
            try:
                # 最初の値を日付に変換してみる
                pd.to_datetime(df[col].iloc[0], errors='raise')
                date_col = col
                break
            except:
                continue

        if date_col is None:
            st.error("日付列が見つかりません。正しい形式のデータをアップロードしてください。")
            st.stop()

        # 日付列を datetime に変換
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')

        for _, row in df.iterrows():
            date = row[date_col]
            if pd.isnull(date):
                continue

            mm = date.month
            month_index = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'

            # 列名リストを取得（列順に基づいて処理する）
column_names = df.columns.tolist()

# 日付列名を特定（例：'年月日' または先頭の列名）
date_col = column_names[0]  # 例: '2024/07/01' など

for _, row in df.iterrows():
    date = pd.to_datetime(row[date_col], errors='coerce')
    if pd.isnull(date):
        continue

    mm = date.month
    month_index = mm if mm >= 4 else mm + 12
    key = 'holiday' if is_holiday(date) else 'weekday'

    # 2列目以降の48個の列を処理
    for i in range(1, 49):  # 1～48列目（0番目は日付）
        if i >= len(column_names):  # 列が足りない場合はスキップ
            continue
        colname = column_names[i]
        val = pd.to_numeric(row[colname], errors='coerce')
        if not pd.isnull(val):
            monthly_data[key][month_index][i - 1] += val  # i-1が0～47のインデックス


    # 雛形読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 平日
    for m in range(4, 16):
        col = chr(64 + m - 1)  # C〜N
        for i in range(48):
            ws[f"{col}{4+i}"] = monthly_data['weekday'][m][i]

    # 休日
    for m in range(4, 16):
        col_index = 17 + (m - 4)  # Q=17列目
        col = chr(64 + col_index if col_index <= 26 else chr(64 + col_index // 26 - 1) + chr(64 + col_index % 26))
        for i in range(48):
            ws[f"{col}{4+i}"] = monthly_data['holiday'][m][i]

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
