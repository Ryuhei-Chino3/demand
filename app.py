import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from copy import deepcopy

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

output_name = st.text_input("出力ファイル名（拡張子不要）", value="", help="例：cats_202406 ※必須")

run_button = st.button("✅ 実行")

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

# アップロードファイル読み込み関数
def read_uploaded(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=5)
        df['Sheet'] = file.name
        return [df]
    else:
        xlsx = pd.ExcelFile(file)
        all_sheets = []
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, header=5)
            df['Sheet'] = sheet_name
            all_sheets.append(df)
        return all_sheets

# 実行ボタン押されたときのみ処理実行
if run_button:
    if not uploaded_files:
        st.warning("ファイルをアップロードしてください。")
        st.stop()

    if output_name.strip() == "":
        st.warning("出力ファイル名を入力してください。")
        st.stop()

    monthly_data = init_monthly_data()
    wb = load_workbook(template_file)
    ws_template: Worksheet = wb["コマ単位集計雛形（送電端）"]

    for file in uploaded_files:
        dataframes = read_uploaded(file)
        for df in dataframes:
            if df.empty:
                continue

            df_columns = df.columns.tolist()

            for _, row in df.iterrows():
                date = pd.to_datetime(row[df_columns[0]], errors='coerce')
                if pd.isnull(date):
                    continue

                mm = date.month
                month_index = mm if mm >= 4 else mm + 12
                key = 'holiday' if is_holiday(date) else 'weekday'

                for i in range(1, 49):  # 1列目から48列目（0は日付）
                    if i >= len(df_columns):
                        continue
                    colname = df_columns[i]
                    val = pd.to_numeric(row[colname], errors='coerce')
                    if not pd.isnull(val):
                        monthly_data[key][month_index][i - 1] += val

            # シート追加：5行目以降をそのまま別シートへ
            month_str = pd.to_datetime(df[df_columns[0]].dropna().iloc[0]).strftime("%Y%m")
            df_with_header = pd.read_excel(file, sheet_name=df['Sheet'].iloc[0], header=None)
            sheet_name = month_str
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws_data = wb.create_sheet(title=sheet_name)
            for r in df_with_header.itertuples(index=False):
                ws_data.append(r)

    # ✅ 平日データ → C〜N列（3〜14列）
    for m in range(4, 16):
        col_idx = m - 1  # 月→列: 4月→3, ..., 翌年3月→14
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]

    # ✅ 休日データ → Q〜AB列（17〜28列）
    for m in range(4, 16):
        col_idx = 17 + (m - 4)  # 修正: 4月→17(Q), ..., 翌年3月→28(AB)
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]


    # 出力ファイル作成
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("変換が完了しました！")
    st.download_button(
        label="📥 処理済みExcelをダウンロード",
        data=output,
        file_name=output_name.strip() + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
