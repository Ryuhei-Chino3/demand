import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

st.title("30分値 平日・休日別 月別合計アプリ")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

output_filename = st.text_input("出力ファイル名（拡張子不要）", value="", help="例: output_202406")
run_button = st.button("実行")

if uploaded_file and output_filename and run_button:
    # 入力ファイルの読み込み
    xls = pd.ExcelFile(uploaded_file)
    monthly_data = {'weekday': {}, 'holiday': {}}
    raw_data_frames = {}

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=4)  # 5行目から読み込み
        raw_data_frames[sheet_name] = df.copy()
        df.columns = df.columns.astype(str)
        
        # 日付と時間の列が存在することを確認
        if '日付' not in df.columns or '時間' not in df.columns or '使用量' not in df.columns:
            st.error(f"{sheet_name} シートに必要な列（日付, 時間, 使用量）が見つかりません")
            st.stop()
        
        df['datetime'] = pd.to_datetime(df['日付'].astype(str) + ' ' + df['時間'].astype(str), errors='coerce')
        df.dropna(subset=['datetime'], inplace=True)

        df['month'] = df['datetime'].dt.month
        df['date'] = df['datetime'].dt.date
        df['dow'] = df['datetime'].dt.dayofweek
        df['is_holiday'] = df['dow'] >= 5  # 土:5, 日:6

        month = df['month'].iloc[0]
        weekday_data = df[df['is_holiday'] == False].groupby('時間')['使用量'].sum().reindex(df['時間'].unique(), fill_value=0)
        holiday_data = df[df['is_holiday'] == True].groupby('時間')['使用量'].sum().reindex(df['時間'].unique(), fill_value=0)

        # 月単位で格納（4〜15月を想定、実データに合わせて拡張可）
        monthly_data['weekday'][month + 0] = weekday_data.tolist()
        monthly_data['holiday'][month + 0] = holiday_data.tolist()

    # 雛形テンプレート読み込み
    template = load_workbook("雛形_伊藤忠.xlsx")
    ws_template = template.active

    # 平日：C〜N列（列番号3〜14）
    for m in range(4, 16):  # 4〜15月
        if m not in monthly_data['weekday']:
            continue
        col_idx = m - 1  # 4月→3(C)
        col_letter = get_column_letter(col_idx + 1)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['weekday'][m][i]

    # 休日：Q〜AB列（列番号17〜30）
    for m in range(4, 16):
        if m not in monthly_data['holiday']:
            continue
        col_idx = 16 + (m - 4)  # 4月→16(Q)
        col_letter = get_column_letter(col_idx + 1)
        for i in range(48):
            ws_template[f"{col_letter}{4+i}"] = monthly_data['holiday'][m][i]

    # 入力ファイルの元データを月ごとに別シートで追加（ヘッダー含む）
    for sheet_name, df in raw_data_frames.items():
        month_str = df['datetime'].dt.strftime('%Y%m').iloc[0]
        ws_month = template.create_sheet(title=month_str)

        # ヘッダー
        for col_idx, col_name in enumerate(df.columns, start=1):
            ws_month.cell(row=1, column=col_idx, value=col_name)

        # データ
        for row_idx, row in enumerate(df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                ws_month.cell(row=row_idx, column=col_idx, value=value)

    # バッファへ保存
    output = BytesIO()
    template.save(output)
    output.seek(0)

    st.success("出力が完了しました。以下からダウンロードしてください。")
    st.download_button(
        label="結果ファイルをダウンロード",
        data=output,
        file_name=f"{output_filename}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
elif uploaded_file and not output_filename and run_button:
    st.warning("出力ファイル名を入力してください。")
