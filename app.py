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

# 曜日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 空の月別データ構造を作成（4月〜翌年3月）
def init_monthly_data():
    return {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)},
        'weekday_days': {month: 0 for month in range(4, 16)},
        'holiday_days': {month: 0 for month in range(4, 16)}
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
    latest_month_map = {}  # 各月ごとに最新のデータフレームを保持

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
            month_key = first_date.year * 100 + first_date.month  # 例：202406

            # 月単位で最新のデータフレームだけを保持
            if (month_key not in latest_month_map) or (first_date > latest_month_map[month_key]["date"]):
                latest_month_map[month_key] = {
                    "df": df,
                    "file": file,
                    "sheet": df['Sheet'].iloc[0],
                    "date": first_date
                }

    # 月ごとに1件ずつ処理
    for m_key, info in latest_month_map.items():
        df = info["df"]
        file = info["file"]
        sheet_name = info["sheet"]

        df_columns = df.columns.tolist()
        used_dates_weekday = set()
        used_dates_holiday = set()

        for _, row in df.iterrows():
            date = pd.to_datetime(row[df_columns[0]], errors='coerce')
            if pd.isnull(date):
                continue

            mm = date.month
            m_idx = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'

            # 日数カウント（同一日で複数行あっても1回だけカウント）
            date_str = date.strftime("%Y-%m-%d")
            if key == 'weekday' and date_str not in used_dates_weekday:
                monthly_data['weekday_days'][m_idx] += 1
                used_dates_weekday.add(date_str)
            elif key == 'holiday' and date_str not in used_dates_holiday:
                monthly_data['holiday_days'][m_idx] += 1
                used_dates_holiday.add(date_str)

            for i in range(1, 49):
                if i >= len(df_columns):
                    continue
                val = pd.to_numeric(row[df_columns[i]], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][m_idx][i - 1] += val

        # シート追加（元データ）
        df_with_header = pd.read_excel(file, sheet_name=sheet_name, header=None)
        output_sheet_name = info["date"].strftime("%Y%m")
        if output_sheet_name in wb.sheetnames:
            del wb[output_sheet_name]
        ws_data = wb.create_sheet(title=output_sheet_name)
        for r in df_with_header.itertuples(index=False):
            ws_data.append(r)

    # ✅ 平日データ → C〜N列（3〜14列） + C57〜N57に日数
    for m in range(4, 16):
        col_idx = m - 1  # 4月→3列目（C）
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]
        ws_template[f"{col_letter}57"] = monthly_data['weekday_days'][m]

    # ✅ 休日データ → Q〜AB列（17〜28列） + Q57〜AB57に日数
    for m in range(4, 16):
        col_idx = 17 + (m - 4)  # 4月→17列目（Q）
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]
        ws_template[f"{col_letter}57"] = monthly_data['holiday_days'][m]

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
