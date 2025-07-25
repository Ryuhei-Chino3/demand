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

# 祝日・休日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 月別集計用データ構造
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)}
    }

# 入力ファイル読み込み（6行目以降）
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

# 変換処理本体
if uploaded_files:
    monthly_data = init_monthly_data()

    for file in uploaded_files:
        df = read_uploaded(file)
        df = df.dropna(how='all')  # 完全空行を削除

        # 列名を付ける（1列目：日付、2列目〜49列目：30分値）
        column_names = ['日時'] + [f"{i}" for i in range(1, 49)]
        if len(df.columns) >= 49:
            df.columns = column_names + list(df.columns[49:])
        else:
            df.columns = column_names[:len(df.columns)]

        for _, row in df.iterrows():
            try:
                date = pd.to_datetime(row['日時'], errors='coerce')
                if pd.isnull(date):
                    continue
            except Exception:
                continue

            mm = date.month
            month_index = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'

            for i in range(1, 49):  # 1〜48列（30分値）
                colname = str(i)
                val = pd.to_numeric(row.get(colname, np.nan), errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

    # 雛形ファイル読み込み
    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 平日列出力（4月〜翌年3月）：列C〜Nに対応（E=6月, F=7月）
    for m in range(4, 16):
        col_index = m - 1  # 4月→C列(3番目) → Excel列インデックス:3,4,...
        col_letter = get_column_letter(col_index)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['weekday'][m][i]

    # 休日列出力（Q〜AB列に対応）
    for m in range(4, 16):
        col_index = 16 + (m - 4)  # 4月→Q列(17), 5月→R列(18)...
        col_letter = get_column_letter(col_index)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['holiday'][m][i]

    # ダウンロード出力
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name="output_koma_format.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
