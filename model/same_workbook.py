"""同一の抽出ラベルExcel内で、UNIT内結線図群と任意の1シートを比較する機能。"""
from __future__ import annotations

import io
from typing import Iterable

import pandas as pd

SUMMARY_SHEET = "Summary"
TOTAL_SHEET = "Total"
LABEL_COLUMN = "ラベル"
COUNT_COLUMN = "個数"
DRAWING_COLUMN = "図番"
TITLE_COLUMN = "タイトル"
UNIT_TITLE_TEXT = "UNIT内結線図"


def normalize_width(value: object) -> str:
    """全角ASCIIと全角スペースだけを半角へ寄せる。"""
    text = "" if pd.isna(value) else str(value)
    converted = []
    for char in text:
        code = ord(char)
        if 0xFF01 <= code <= 0xFF5E:
            converted.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            converted.append(" ")
        else:
            converted.append(char)
    return "".join(converted)


def _text(value: object) -> str:
    return "" if pd.isna(value) else str(value)


def _require_columns(df: pd.DataFrame, sheet_name: str, columns: Iterable[str]) -> None:
    missing = [column for column in columns if column not in df.columns]
    if missing:
        raise ValueError(f"'{sheet_name}' シートに必要な列がありません: {', '.join(missing)}")


def list_label_sheets(file_bytes: bytes) -> list[str]:
    """Bとして選択できる、ラベル・個数列をもつ個別図面シート名を返す。"""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    candidates = []
    for sheet_name in xls.sheet_names:
        if sheet_name in (SUMMARY_SHEET, TOTAL_SHEET):
            continue
        df = xls.parse(sheet_name, nrows=0)
        if LABEL_COLUMN in df.columns and COUNT_COLUMN in df.columns:
            candidates.append(sheet_name)
    return candidates


def compare_within_workbook(file_bytes: bytes, b_sheet_name: str) -> dict:
    """同一Excel内のA（UNIT内結線図群）とB（指定シート）を比較する。

    A選択時のみタイトルの全角・半角を同一視する。ラベルは元の表記のまま比較し、
    元の依頼どおりにAの重複除外リストを作る。
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if SUMMARY_SHEET not in xls.sheet_names:
        raise ValueError(f"'{SUMMARY_SHEET}' シートが見つかりません")
    if b_sheet_name not in xls.sheet_names:
        raise ValueError(f"指定したシートが見つかりません: {b_sheet_name}")
    if b_sheet_name in (SUMMARY_SHEET, TOTAL_SHEET):
        raise ValueError("Bには個別図面のラベルシートを指定してください")

    summary_df = xls.parse(SUMMARY_SHEET)
    _require_columns(summary_df, SUMMARY_SHEET, (DRAWING_COLUMN, TITLE_COLUMN))
    selected_drawings = []
    for _, row in summary_df.iterrows():
        drawing_no = _text(row[DRAWING_COLUMN]).strip()
        title = _text(row[TITLE_COLUMN])
        if drawing_no and UNIT_TITLE_TEXT in normalize_width(title).upper():
            selected_drawings.append(drawing_no)

    if not selected_drawings:
        raise ValueError("タイトルに 'UNIT内結線図' を含む図番が Summary シートにありません")

    a_labels: dict[str, dict] = {}
    missing_sheets = []
    for drawing_no in selected_drawings:
        if drawing_no not in xls.sheet_names:
            missing_sheets.append(drawing_no)
            continue
        df = xls.parse(drawing_no)
        _require_columns(df, drawing_no, (LABEL_COLUMN, COUNT_COLUMN))
        for label, count in zip(df[LABEL_COLUMN], df[COUNT_COLUMN]):
            if pd.isna(label):
                continue
            label_text = _text(label)
            if not label_text:
                continue
            entry = a_labels.setdefault(label_text, {"count": 0, "drawings": set()})
            entry["count"] += int(count) if pd.notna(count) else 0
            entry["drawings"].add(drawing_no)

    if missing_sheets:
        raise ValueError(
            "Summaryの図番に対応するシートがありません: " + ", ".join(missing_sheets))

    b_df = xls.parse(b_sheet_name)
    _require_columns(b_df, b_sheet_name, (LABEL_COLUMN, COUNT_COLUMN))
    b_labels: dict[str, int] = {}
    for label, count in zip(b_df[LABEL_COLUMN], b_df[COUNT_COLUMN]):
        if pd.isna(label):
            continue
        label_text = _text(label)
        if not label_text:
            continue
        b_labels[label_text] = b_labels.get(label_text, 0) + (int(count) if pd.notna(count) else 0)

    labels = sorted(set(a_labels) | set(b_labels))
    comparison_rows = []
    for label in labels:
        in_a, in_b = label in a_labels, label in b_labels
        status = "両方" if in_a and in_b else "A のみ" if in_a else "B のみ"
        a_entry = a_labels.get(label)
        comparison_rows.append({
            "ラベル": label,
            "A 合計出現数": a_entry["count"] if a_entry else 0,
            "A 図番数": len(a_entry["drawings"]) if a_entry else 0,
            "B 出現数": b_labels.get(label, 0),
            "A 図番": ", ".join(sorted(a_entry["drawings"])) if a_entry else "",
            "比較結果": status,
        })

    result_df = pd.DataFrame(comparison_rows)
    return {
        "b_sheet_name": b_sheet_name,
        "comparison": result_df,
        "summary": {
            "A 対象図番数": len(selected_drawings),
            "A ユニークラベル数": len(a_labels),
            "B ユニークラベル数": len(b_labels),
            "両方": int((result_df["比較結果"] == "両方").sum()),
            "A のみ": int((result_df["比較結果"] == "A のみ").sum()),
            "B のみ": int((result_df["比較結果"] == "B のみ").sum()),
        },
    }
