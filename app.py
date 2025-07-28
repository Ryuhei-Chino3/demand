import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO
import calendar
import os

st.set_page_config(page_title="30分値集計アプリ", layout="wide")
st.title("30分値 平日・休日 月別集計アプリ")

uploaded_files = st.file_uploader("30分値のExcelまたはCSVファイルをアップロード", type=["xlsx", "csv"], accept_multiple_files=True)
output_filename = st.text_input("出力ファイル名（拡張子 .xlsx は不要）", value="")
execute = st.button("変換を実行")

if execute:
    if not uploaded_files:
        st.warning("ファイルをアップロードしてください。")
    elif not output_filename.strip():
        st.warning("出力ファイル名を入力してください。")
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "集計結果"

        # ヘッダー行
        ws.append(["区分", "契約容量(kW)"] + [f"{month}月 平日" for month in range(1, 13)] + [f"{month}月 休日" for month in range(1, 13)])

        for uploaded_file in uploaded_files:
            # ファイル読み込み
            if uploaded_file.name.endswith(".csv"):
                try:
                    df_raw = pd.read_csv(uploaded_file, encoding="utf-8", header=None)
                except UnicodeDecodeError:
                    df_raw = pd.read_csv(uploaded_file, encoding="shift_jis", header=None)
            else:
                df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

            # 5行目以降をデータとして読み込み
            df = df_raw.iloc[4:].copy()
            df.columns = df_raw.iloc[4].tolist()  # 5行目をヘッダーとして使用
            df = df[1:].reset_index(drop=True)

            # 列名の正規化
            cols = df.columns.tolist()
            col_map = {}
            for col in cols:
                if '日付' in str(col):
                    col_map[col] = 'date'
                elif '時間' in str(col):
                    col_map[col] = 'time'
                elif '使用量' in str(col) or '使用電力量' in str(col):
                    col_map[col] = 'usage'
            df = df.rename(columns=col_map)

            # 必須列チェック
            if not {'date', 'time', 'usage'}.issubset(df.columns):
                st.error(f"{uploaded_file.name} に必要な列（日付, 時間, 使用量）が見つかりません。")
                st.stop()

            # 日付・時間をdatetimeに
            df['datetime'] = pd.to_datetime(df['date'].astype(str) + ' ' + df['time'].astype(str), errors='coerce')
            df['usage'] = pd.to_numeric(df['usage'], errors='coerce')
            df = df.dropna(subset=['datetime', 'usage'])

            df['month'] = df['datetime'].dt.month
            df['weekday'] = df['datetime'].dt.weekday
            df['is_holiday'] = df['weekday'] >= 5

            # 月ごとの平日・休日集計
            monthly_data = {'weekday': {}, 'holiday': {}}
            for m in range(1, 13):
                df_m = df[df['month'] == m]
                for kind, label in [(False, 'weekday'), (True, 'holiday')]:
                    usage_sum = df_m[df_m['is_holiday'] == kind].groupby(df_m['datetime'].dt.time)['usage'].sum()
                    if not usage_sum.empty:
                        monthly_data[label][m] = usage_sum.reindex(pd.date_range("00:00", "23:30", freq="30min").time, fill_value=0).tolist()
                    else:
                        monthly_data[label][m] = [0]*48

            # 出力行作成
            name = os.path.splitext(uploaded_file.name)[0]
            contract_capacity = 18  # 仮設定、必要に応じて取得方法を変える
            row = [name, contract_capacity]
            for m in range(1, 13):
                row.append(sum(monthly_data['weekday'].get(m, [0]*48)))
            for m in range(1, 13):
                row.append(sum(monthly_data['holiday'].get(m, [0]*48)))
            ws.append(row)

            # 入力シートもコピー
            month_label = df['datetime'].dt.strftime('%Y%m').iloc[0]
            sheet = wb.create_sheet(title=month_label)
            for i, row in enumerate(df_raw.values.tolist()):
                sheet.append(row)

        # ダウンロード用にバッファへ保存
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("変換が完了しました！")
        st.download_button("集計結果をダウンロード", data=output, file_name=f"{output_filename.strip()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
