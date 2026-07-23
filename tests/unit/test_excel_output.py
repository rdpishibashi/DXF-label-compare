import io
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from model.compare_labels import compare_labels, summarize, compare_labels_by_region, build_region_summary_rows
from model.excel_output import create_compare_excel_output, create_region_compare_excel_output


def test_create_compare_excel_output_sheets_and_columns():
    df = compare_labels({'CN1': 2, 'R10': 1}, {'CN1': 5, 'X1': 1})
    summary = summarize(df, 'A.xlsx', 'B.xlsx')
    xlsx_bytes = create_compare_excel_output(df, summary)

    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    assert xls.sheet_names == ['サマリー', '差分']
    diff = xls.parse('差分')
    assert list(diff.columns) == ['ラベル', '区分', 'A個数', 'B個数']
    assert len(diff) == len(df)
    # A のみのラベル(R10)はB個数が空欄(NaN)
    r10 = diff[diff['ラベル'] == 'R10'].iloc[0]
    assert pd.isna(r10['B個数'])


def test_create_region_compare_excel_output_sheets_and_columns():
    a_by_region = {'R1': {'CN1': 2, 'R10': 1}}
    b_by_region = {'R1': {'CN1': 5, 'X1': 1}}
    diff_df, metrics = compare_labels_by_region(a_by_region, b_by_region, ['R1'])
    summary_rows = build_region_summary_rows(metrics, 'A.xlsx', 'B.xlsx', b_filter_mode='全部')

    xlsx_bytes = create_region_compare_excel_output(diff_df, summary_rows)
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    assert xls.sheet_names == ['サマリー', '差分']

    diff = xls.parse('差分')
    assert list(diff.columns) == ['領域名', 'ラベル', '区分', 'A個数', 'B個数']
    # 同じ領域名が連続する行は2行目以降が空欄になる（先頭行だけ'R1'）
    assert diff.iloc[0]['領域名'] == 'R1'
    assert diff['領域名'].iloc[1:].isna().all()

    summ = xls.parse('サマリー')
    assert list(summ.columns) == ['領域名', '項目', '値']
    assert summ.iloc[0]['項目'] == 'A ファイル名'
    assert pd.isna(summ.iloc[0]['領域名'])  # ヘッダー行は領域名が空欄
    assert summ.iloc[3]['領域名'] == 'R1'  # 領域ブロックの先頭行だけ領域名が入る
    assert summ['領域名'].iloc[4:8].isna().all()
