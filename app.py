import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime
from io import BytesIO
import calendar
import os

st.set_page_config(page_title="フォーマット変換アプリ", layout="wide")
st.title("伊藤忠フォーマット変換アプリ")

uploaded_files = st.file_uploader("30分値のExcelまたはCSVファイルをアップロード", type=["xlsx", "csv"], accept_multiple_files=True)
output_filename = st.text_input("出力ファイル名（拡張子 .xlsx は不要）", value="")
execute = st.button("変換を実行")

if execute:
    if not uploaded_files:
        st.warning("ファイルをアップロードしてください。")
    elif not output_filename.strip():
        st.warning("出力ファイル名を入力してください。")
    else:
        # 雛形ファイルを読み込み
        template_wb = openpyxl.load_workbook('雛形_伊藤忠.xlsx')
        template_ws = template_wb['コマ単位集計雛形（送電端）']
        
        # 新しいワークブックを作成
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "コマ単位集計雛形（送電端）"
        
        # 雛形の構造と書式を完全にコピー
        for row in template_ws.iter_rows():
            for cell in row:
                new_cell = ws[cell.coordinate]
                new_cell.value = cell.value
                
                # 書式をコピー
                if cell.font:
                    new_cell.font = cell.font
                if cell.border:
                    new_cell.border = cell.border
                if cell.fill:
                    new_cell.fill = cell.fill
                if cell.number_format:
                    new_cell.number_format = cell.number_format
                if cell.alignment:
                    new_cell.alignment = cell.alignment
        
        # 列幅と行高をコピー
        for col in range(1, template_ws.max_column + 1):
            col_letter = get_column_letter(col)
            if template_ws.column_dimensions[col_letter].width:
                ws.column_dimensions[col_letter].width = template_ws.column_dimensions[col_letter].width
        
        for row in range(1, template_ws.max_row + 1):
            if template_ws.row_dimensions[row].height:
                ws.row_dimensions[row].height = template_ws.row_dimensions[row].height
        
        # 30分間隔の時間リストを作成
        time_slots = []
        for hour in range(24):
            for minute in [0, 30]:
                start_time = f"{hour:02d}:{minute:02d}"
                if minute == 30:
                    end_hour = hour
                    end_minute = 0
                else:
                    end_hour = hour
                    end_minute = 30
                end_time = f"{end_hour:02d}:{end_minute:02d}"
                time_slots.append(f"{start_time}-{end_time}")

        for uploaded_file in uploaded_files:
            # ファイル読み込み
            if uploaded_file.name.endswith(".csv"):
                try:
                    df_raw = pd.read_csv(uploaded_file, encoding="utf-8", header=None)
                except UnicodeDecodeError:
                    df_raw = pd.read_csv(uploaded_file, encoding="shift_jis", header=None)
            else:
                df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)

            # データ形式を判定
            # 5行目以降をデータとして読み込み
            df = df_raw.iloc[4:].copy()
            df.columns = df_raw.iloc[4].tolist()  # 5行目をヘッダーとして使用
            df = df[1:].reset_index(drop=True)

            # データ形式を判定（横型か縦型か）
            is_horizontal_format = False
            if len(df.columns) > 10:  # 横型の場合は時間列が多数ある
                time_columns = [col for col in df.columns if isinstance(col, str) and ':' in col]
                if len(time_columns) >= 48:  # 30分値なら48列
                    is_horizontal_format = True

            if is_horizontal_format:
                # 横型データの処理
                st.info(f"{uploaded_file.name}: 横型データ形式を検出しました")
                
                # 日付列を取得
                date_col = None
                for col in df.columns:
                    if '年月日' in str(col) or '日付' in str(col):
                        date_col = col
                        break
                
                if date_col is None:
                    st.error(f"{uploaded_file.name} に日付列が見つかりません。")
                    continue

                # 時間列を取得（数値データのみ）
                time_columns = []
                for col in df.columns:
                    if isinstance(col, str) and ':' in col:
                        time_columns.append(col)

                if not time_columns:
                    st.error(f"{uploaded_file.name} に時間列が見つかりません。")
                    continue

                # データを縦型に変換
                df_melted = df.melt(id_vars=[date_col], value_vars=time_columns, 
                                   var_name='time', value_name='usage')
                
                # 日付をdatetimeに変換
                df_melted['date'] = pd.to_datetime(df_melted[date_col], errors='coerce')
                df_melted['usage'] = pd.to_numeric(df_melted['usage'], errors='coerce')
                df_melted = df_melted.dropna(subset=['date', 'usage'])

                # datetimeを作成
                df_melted['datetime'] = pd.to_datetime(
                    df_melted['date'].dt.strftime('%Y-%m-%d') + ' ' + df_melted['time'].astype(str), 
                    errors='coerce'
                )
                
                df = df_melted[['datetime', 'usage']].copy()

            else:
                # 縦型データの処理（元のコード）
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
                    continue

                # 日付・時間をdatetimeに
                df['datetime'] = pd.to_datetime(df['date'].astype(str) + ' ' + df['time'].astype(str), errors='coerce')
                df['usage'] = pd.to_numeric(df['usage'], errors='coerce')
                df = df.dropna(subset=['datetime', 'usage'])

            # 月・曜日情報を追加
            df['month'] = df['datetime'].dt.month
            df['weekday'] = df['datetime'].dt.weekday
            df['is_holiday'] = df['weekday'] >= 5
            df['time_slot'] = df['datetime'].dt.strftime('%H:%M')

            # 月ごとの平日・休日集計（30分単位）
            monthly_data = {'weekday': {}, 'holiday': {}}
            
            for m in range(1, 13):
                df_m = df[df['month'] == m]
                
                for kind, label in [(False, 'weekday'), (True, 'holiday')]:
                    df_kind = df_m[df_m['is_holiday'] == kind]
                    
                    if not df_kind.empty:
                        # 30分単位で集計
                        usage_by_time = df_kind.groupby('time_slot')['usage'].sum()
                        
                        # 48個の30分スロットに値を設定
                        time_data = {}
                        for time_slot in time_slots:
                            start_time = time_slot.split('-')[0]
                            if start_time in usage_by_time.index:
                                time_data[time_slot] = usage_by_time[start_time]
                            else:
                                time_data[time_slot] = 0
                        
                        monthly_data[label][m] = time_data
                    else:
                        monthly_data[label][m] = {time_slot: 0 for time_slot in time_slots}

            # 伊藤忠フォーマットにデータを挿入
            # 平日データを挿入
            for i, time_slot in enumerate(time_slots):
                row_idx = i + 3  # 3行目から開始
                
                for month in range(1, 13):
                    col_idx = month + 1  # 2列目から開始（1列目は時間）
                    if month in monthly_data['weekday']:
                        value = monthly_data['weekday'][month].get(time_slot, 0)
                        ws.cell(row=row_idx, column=col_idx, value=value)
            
            # 休日データを挿入（16列目から開始）
            for i, time_slot in enumerate(time_slots):
                row_idx = i + 3  # 3行目から開始
                
                for month in range(1, 13):
                    col_idx = month + 15  # 16列目から開始
                    if month in monthly_data['holiday']:
                        value = monthly_data['holiday'][month].get(time_slot, 0)
                        ws.cell(row=row_idx, column=col_idx, value=value)

            # 入力ファイルのデータを別シートに保存
            month_label = df['datetime'].dt.strftime('%Y%m').iloc[0]
            sheet_name = f"{os.path.splitext(uploaded_file.name)[0]}_{month_label}"
            sheet = wb.create_sheet(title=sheet_name)
            
            # 元のデータをそのままコピー
            for i, row in enumerate(df_raw.values.tolist()):
                for j, value in enumerate(row):
                    sheet.cell(row=i+1, column=j+1, value=value)

        # ダウンロード用にバッファへ保存
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("変換が完了しました！")
        st.download_button("集計結果をダウンロード", data=output, file_name=f"{output_filename.strip()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
