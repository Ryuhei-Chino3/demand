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

st.title("伊藤忠フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

output_name = st.text_input("出力ファイル名（拡張子不要）", value="", help="例：cats_202406 ※必須")

run_button = st.button("✅ 実行")

template_file = "雛形_伊藤忠.xlsx"

# 休日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 月別初期データ
def init_monthly_data():
    return {
        'weekday': {m: [0]*48 for m in range(4, 16)},
        'holiday': {m: [0]*48 for m in range(4, 16)},
        'weekday_days': {m: 0 for m in range(4, 16)},
        'holiday_days': {m: 0 for m in range(4, 16)}
    }

# ファイル読み込み（CSVまたはXLSX）
def read_uploaded(file):
    if file.name.endswith('.csv'):
        df = pd.read_csv(file, header=5)
        df['Sheet'] = file.name
        return [df]
    else:
        xlsx = pd.ExcelFile(file)
        return [pd.read_excel(xlsx, sheet_name=sheet, header=5).assign(Sheet=sheet) for sheet in xlsx.sheet_names]

if run_button:
    if not uploaded_files:
        st.warning("ファイルをアップロードしてください。")
        st.stop()

    if output_name.strip() == "":
        st.warning("出力ファイル名を入力してください。")
        st.stop()

    monthly_data = init_monthly_data()
    latest_month_map = {}

    wb = load_workbook(template_file)
    ws_template: Worksheet = wb["コマ単位集計雛形（送電端）"]

    for file in uploaded_files:
        dataframes = read_uploaded(file)
        for df in dataframes:
            if df.empty:
                continue

            df_columns = df.columns.tolist()
            dates = pd.to_datetime(df[df_columns[0]], errors='coerce')
            valid_dates = dates.dropna()
            if valid_dates.empty:
                continue

            first_date = valid_dates.iloc[0]
            month_key = first_date.year * 100 + first_date.month  # 年月で一意に判定

            if (month_key not in latest_month_map) or (first_date > latest_month_map[month_key]["date"]):
                latest_month_map[month_key] = {
                    "df": df,
                    "file": file,
                    "sheet": df['Sheet'].iloc[0],
                    "date": first_date
                }

    for m_key, info in latest_month_map.items():
        df = info["df"]
        file = info["file"]
        sheet_name = info["sheet"]

        df_columns = df.columns.tolist()
        used_dates = {'weekday': set(), 'holiday': set()}

        for _, row in df.iterrows():
            date = pd.to_datetime(row[df_columns[0]], errors='coerce')
            if pd.isnull(date):
                continue

            mm = date.month
            month_index = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'

            date_str = date.strftime("%Y-%m-%d")
            if date_str not in used_dates[key]:
                monthly_data[key + '_days'][month_index] += 1
                used_dates[key].add(date_str)

            for i in range(1, 49):
                if i >= len(df_columns):
                    continue
                val = pd.to_numeric(row[df_columns[i]], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

        # 元データシート追加
        df_with_header = pd.read_excel(file, sheet_name=sheet_name, header=None)
        sheet_title = info["date"].strftime("%Y%m")
        if sheet_title in wb.sheetnames:
            del wb[sheet_title]
        ws_data = wb.create_sheet(title=sheet_title)
        for r in df_with_header.itertuples(index=False):
            ws_data.append(r)

    # ✅ 平日：C〜N列 & C57〜N57（日数）
    for m in range(4, 16):
        col_idx = m - 1  # C=3（4月）
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]
        ws_template[f"{col_letter}57"] = monthly_data['weekday_days'][m]

    # ✅ 休日：Q〜AB列 & Q57〜AB57（日数）
    for m in range(4, 16):
        col_idx = 17 + (m - 4)  # Q=17（4月）
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]
        ws_template[f"{col_letter}57"] = monthly_data['holiday_days'][m]

    # 出力
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
