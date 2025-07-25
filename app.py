import streamlit as st
import pandas as pd
import numpy as np
import jpholiday
import datetime
import calendar
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import deepcopy

st.title("30分値 → 雛形フォーマット変換アプリ")

uploaded_files = st.file_uploader("ファイルをアップロード（複数可）", type=['xlsx', 'csv'], accept_multiple_files=True)

template_file = "雛形_伊藤忠.xlsx"

# 曜日判定
def is_holiday(date):
    return date.weekday() >= 5 or jpholiday.is_holiday(date)

# 空の月別データ構造を作成（4月〜翌年3月 → 4〜15で扱う）
def init_monthly_data():
    return {
