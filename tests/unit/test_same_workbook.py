import io

import pandas as pd

from utils.same_workbook import compare_within_workbook, list_label_sheets


def _workbook_bytes():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame([
            {"図番": "EE001", "タイトル": "ＵＮＩＴ内結線図", "サブタイトル": "FIRST"},
            {"図番": "EE002", "タイトル": "UNIT内結線図", "サブタイトル": "SECOND"},
            {"図番": "EE003", "タイトル": "部品図", "サブタイトル": "OTHER"},
        ]).to_excel(writer, sheet_name="Summary", index=False)
        pd.DataFrame([{"ラベル": "A", "個数": 2}, {"ラベル": "C", "個数": 1}]).to_excel(
            writer, sheet_name="EE001", index=False)
        pd.DataFrame([{"ラベル": "A", "個数": 3}, {"ラベル": "B", "個数": 1}]).to_excel(
            writer, sheet_name="EE002", index=False)
        pd.DataFrame([{"ラベル": "Z", "個数": 1}]).to_excel(writer, sheet_name="EE003", index=False)
        pd.DataFrame([{"ラベル": "A", "個数": 5}, {"ラベル": "D", "個数": 1}]).to_excel(
            writer, sheet_name="EE999", index=False)
        pd.DataFrame([{"ラベル": "A", "個数": 5, "図番": "EE001"}]).to_excel(
            writer, sheet_name="Total", index=False)
    return output.getvalue()


def test_list_label_sheets_excludes_summary_and_total():
    assert list_label_sheets(_workbook_bytes()) == ["EE001", "EE002", "EE003", "EE999"]


def test_compare_within_workbook_filters_unit_titles_and_aggregates_labels():
    result = compare_within_workbook(_workbook_bytes(), "EE999")

    assert result["summary"] == {
        "A 対象図番数": 2,
        "A ユニークラベル数": 3,
        "B ユニークラベル数": 2,
        "共通ラベル数": 1,
        "A のみ": 2,
        "B のみ": 1,
    }
    assert result["a_labels"] == [
        {"ラベル": "A", "合計出現数": 5, "図番数": 2, "図番": "EE001, EE002"},
        {"ラベル": "B", "合計出現数": 1, "図番数": 1, "図番": "EE002"},
        {"ラベル": "C", "合計出現数": 1, "図番数": 1, "図番": "EE001"},
    ]
    assert list(result["comparison"]["比較結果"]) == ["共通", "Aのみ", "Aのみ", "Bのみ"]
