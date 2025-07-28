import streamlit as st
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from io import BytesIO
import holidays

# --- 日本の祝日判定 ---
jp_holidays = holidays.Japan()
def is_holiday(date):
    return date.weekday() >= 5 or date in jp_holidays

# --- メイン処理関数 ---
def process_files(uploaded_files):
    monthly_data = {
        'weekday': {month: [0]*48 for month in range(4, 16)},
        'holiday': {month: [0]*48 for month in range(4, 16)},
        'weekday_days': {month: 0 for month in range(4, 16)},
        'holiday_days': {month: 0 for month in range(4, 16)}
    }

    latest_month_map = {}

    # --- ファイル毎に読み取り、最新の月データだけ採用 ---
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith('.xlsx'):
            xls = pd.ExcelFile(uploaded_file)
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                if df.empty or df.shape[1] < 2:
                    continue
                try:
                    first_date = pd.to_datetime(df.iloc[0, 0], errors='coerce')
                    if pd.isnull(first_date):
                        continue
                    mm = first_date.month
                    yy = first_date.year
                    month_key = mm if mm >= 4 else mm + 12
                    ym_key = f"{yy}-{mm:02d}"
                    if (month_key not in latest_month_map) or (ym_key > latest_month_map[month_key]['ym']):
                        latest_month_map[month_key] = {'df': df, 'ym': ym_key, 'file': uploaded_file.name, 'sheet': sheet_name}
                except:
                    continue

    # --- 日付の重複排除マップ（月別・平日/休日別） ---
    used_dates_map = {
        'weekday': {month: set() for month in range(4, 16)},
        'holiday': {month: set() for month in range(4, 16)}
    }

    # --- 最新月データに対して集計 ---
    for m_key, info in latest_month_map.items():
        df = info['df']
        df_columns = df.columns.tolist()

        for _, row in df.iterrows():
            date = pd.to_datetime(row[df_columns[0]], errors='coerce')
            if pd.isnull(date):
                continue
            mm = date.month
            month_index = mm if mm >= 4 else mm + 12
            key = 'holiday' if is_holiday(date) else 'weekday'
            date_str = date.strftime("%Y-%m-%d")

            # ✅ 正しい日数カウント（月別・平日/休日別）
            if date_str not in used_dates_map[key][month_index]:
                monthly_data[key + '_days'][month_index] += 1
                used_dates_map[key][month_index].add(date_str)

            for i in range(1, 49):
                if i >= len(df_columns):
                    continue
                colname = df_columns[i]
                val = pd.to_numeric(row[colname], errors='coerce')
                if not pd.isnull(val):
                    monthly_data[key][month_index][i - 1] += val

    # --- 出力Excel作成 ---
    template_path = "template.xlsx"
    wb = load_workbook(template_path)
    ws_template = wb.active

    # ✅ 平日データ → C〜N列（3〜14列）
    for m in range(4, 16):
        col_letter = get_column_letter(m - 1)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]
        ws_template[f"{col_letter}57"] = monthly_data['weekday_days'][m]  # ✅ 日数

    # ✅ 休日データ → Q〜AB列（17〜30列）
    for m in range(4, 16):
        col_idx = 17 + (m - 4)  # 4月→17(Q), ..., 翌年3月→30(AB)
        col_letter = get_column_letter(col_idx)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]
        ws_template[f"{col_letter}57"] = monthly_data['holiday_days'][m]  # ✅ 日数

    # --- 保存 ---
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit UI ---
st.set_page_config(page_title="伊藤忠フォーマット変換アプリ", layout="centered")
st.title("伊藤忠フォーマット変換アプリ")
st.markdown("### 30分値ファイル（CSV/XLSX）をアップロードしてください")

uploaded_files = st.file_uploader("複数ファイルを選択可能", type=['xlsx', 'csv'], accept_multiple_files=True)

if uploaded_files:
    output = process_files(uploaded_files)
    st.success("変換が完了しました。下記ボタンからダウンロードできます。")
    st.download_button("変換後ファイルをダウンロード", output, file_name="output_format.xlsx")
