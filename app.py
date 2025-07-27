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

if uploaded_files and output_filename:
    monthly_data = init_monthly_data()
    data_sheets = {}  # シート名: DataFrame

    for file in uploaded_files:
        df = read_uploaded(file)

        # 最初の有効な日付を取得して年月シート名作成
        df_dates = pd.to_datetime(df[0], errors='coerce')
        valid_dates = df_dates.dropna()
        if not valid_dates.empty:
            month_str = valid_dates.iloc[0].strftime("%Y%m")
        else:
            month_str = "データ不明"

        data_sheets[month_str] = df.copy()

        for _, row in df.iterrows():
            try:
                date = pd.to_datetime(row[0], errors='coerce')
                if pd.isnull(date):
                    continue

                mm = date.month
                month_index = mm if mm >= 4 else mm + 12
                key = 'holiday' if is_holiday(date) else 'weekday'

                for i in range(48):
                    val = pd.to_numeric(row[i + 1], errors='coerce')
                    if not pd.isnull(val):
                        monthly_data[key][month_index][i] += val
            except Exception:
                continue

    wb = load_workbook(template_file)
    ws = wb["コマ単位集計雛形（送電端）"]

    # 平日（6月→E列, 7月→F列）
    for m in range(6, 8):
        col_index = 4 + (m - 6)
        col_letter = get_column_letter(col_index + 1)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['weekday'][m][i]

    # 休日（6月→S列, 7月→T列）
    for m in range(6, 8):
        col_index = 18 + (m - 6)
        col_letter = get_column_letter(col_index + 1)
        for i in range(48):
            ws[f"{col_letter}{4 + i}"] = monthly_data['holiday'][m][i]

    # 入力元データを別シートに追加
    for sheet_name, data_df in data_sheets.items():
        ws_data = wb.create_sheet(title=sheet_name)
        for r_idx, row in data_df.iterrows():
            for c_idx, val in enumerate(row):
                ws_data.cell(row=r_idx + 1, column=c_idx + 1, value=val)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name=f"{output_filename}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
