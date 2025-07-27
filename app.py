import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

output_filename = st.text_input("出力ファイル名（拡張子は不要）", value="", help="必須項目です。例: 202406_キャッツアイ")

template_file = "雛形_伊藤忠.xlsx"

# 曜日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 空の月別データ構造
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

# ファイル読み込み（5行目以降をDataFrameに）
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

if uploaded_files and output_filename.strip():
    monthly_data = init_monthly_data()

    # 雛形読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

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

        # 入力元データを別シートに出力（5行目以降）
        xlsx = pd.ExcelFile(file)
        for sheet_name in xlsx.sheet_names:
            df_sheet = pd.read_excel(xlsx, sheet_name=sheet_name, header=4)  # 5行目＝index=4から読み込む
            # シート名生成
            first_date = pd.to_datetime(df_sheet.iloc[0, 0], errors='coerce')
            if pd.isnull(first_date):
                continue
            ym_str = first_date.strftime('%Y%m')
            ws_data = wb.create_sheet(title=ym_str)

            for row in dataframe_to_rows(df_sheet, index=False, header=True):
                ws_data.append(row)

    # 平日エリア：E列〜P列（4〜15月 → E〜P）
    for m in range(4, 16):
        col_offset = m - 4  # 0〜11
        col_base = ord('E') + col_offset
        col_letter = chr(col_base)
        for i in range(48):
            ws[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]

    # 休日エリア：S列〜AD列（4〜15月 → S〜AD）
    for m in range(4, 16):
        col_offset = m - 4
        col_base = ord('S') + col_offset
        col_letter = chr(col_base)
        for i in range(48):
            ws[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]

    # 出力
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name=output_filename.strip() + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    if uploaded_files and not output_filename.strip():
        st.warning("出力ファイル名を入力してください。")
