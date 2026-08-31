import io

import pandas as pd

from model.same_workbook_output import create_same_workbook_output


def _result(rows):
    df = pd.DataFrame(rows, columns=[
        "ラベル", "A 合計出現数", "A 図番数", "B 出現数", "A 図番", "比較結果",
    ])
    return {
        "summary": {"A 対象図番数": 1, "A ユニークラベル数": len(df), "B ユニークラベル数": len(df)},
        "comparison": df,
    }


def test_create_same_workbook_output_sheets_and_columns():
    result = _result([
        {"ラベル": "A", "A 合計出現数": 2, "A 図番数": 1, "B 出現数": 2, "A 図番": "EE001", "比較結果": "両方"},
        {"ラベル": "B", "A 合計出現数": 1, "A 図番数": 1, "B 出現数": 0, "A 図番": "EE001", "比較結果": "A のみ"},
    ])
    xlsx_bytes = create_same_workbook_output(result)
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    assert xls.sheet_names == ["サマリー", "差分"]
    diff = xls.parse("差分")
    assert list(diff.columns) == [
        "ラベル", "A 合計出現数", "A 図番数", "B 出現数", "A 図番", "比較結果",
    ]
    assert len(diff) == 2


def test_create_same_workbook_output_handles_both_with_mismatched_counts():
    # 比較結果=両方 だが A 合計出現数 != B 出現数（黄で色分けされる想定の行）。
    # 整数値（pd.NA を経由しない同一ワークブック比較特有の型）でも
    # row_style() が例外なく動作することを確認する。
    result = _result([
        {"ラベル": "A", "A 合計出現数": 5, "A 図番数": 2, "B 出現数": 3, "A 図番": "EE001, EE002", "比較結果": "両方"},
    ])
    xlsx_bytes = create_same_workbook_output(result)
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    diff = xls.parse("差分")
    assert diff.iloc[0]["A 合計出現数"] == 5
    assert diff.iloc[0]["B 出現数"] == 3
    assert diff.iloc[0]["比較結果"] == "両方"


def test_create_same_workbook_output_empty_comparison():
    result = _result([])
    xlsx_bytes = create_same_workbook_output(result)
    xls = pd.ExcelFile(io.BytesIO(xlsx_bytes))
    diff = xls.parse("差分")
    assert len(diff) == 0
