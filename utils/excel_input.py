"""DXF-extract-labels 出力Excelの Total/Summary シート読み込み。"""
import io

import pandas as pd

from utils.compare_labels import normalize_label

REQUIRED_TOTAL_SHEET = 'Total'
REQUIRED_TOTAL_COLUMNS = ('ラベル', '個数')
REQUIRED_TOTAL_COLUMNS_WITH_GZUBAN = ('ラベル', '個数', '図番')
REQUIRED_SUMMARY_SHEET = 'Summary'
REQUIRED_SUMMARY_COLUMNS = ('図番', 'タイトル')


def _parse_total_sheet(file_bytes: bytes, required_columns) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if REQUIRED_TOTAL_SHEET not in xls.sheet_names:
        raise ValueError(f"'{REQUIRED_TOTAL_SHEET}' シートが見つかりません")
    df = xls.parse(REQUIRED_TOTAL_SHEET)
    missing = [c for c in required_columns if c not in df.columns]
    if missing:
        raise ValueError(
            f"{REQUIRED_TOTAL_SHEET} シートに必要な列がありません: {', '.join(missing)}")
    return df


def load_total_labels(file_bytes: bytes) -> dict:
    """Excel の Total シートを読み、正規化ラベル → 合計個数 の dict を返す。

    Total シートが無い、または『ラベル』『個数』列が無い場合は ValueError。
    正規化（全角→半角）で複数の元ラベルが同一キーになった場合は個数を合算する。
    """
    df = _parse_total_sheet(file_bytes, REQUIRED_TOTAL_COLUMNS)
    agg: dict = {}
    for lbl, cnt in zip(df['ラベル'], df['個数']):
        if pd.isna(lbl):
            continue
        key = normalize_label(str(lbl))
        agg[key] = agg.get(key, 0) + int(cnt)
    return agg


def load_total_rows(file_bytes: bytes) -> list:
    """Excel の Total シートを (正規化ラベル, 個数, 図番リスト) のタプルのリストで返す。

    図番 は Total シートの『図番』列（カンマ区切り）を分割・正規化したもの。
    `utils.drawing_filter` の図番フィルタと組み合わせて使う（現状は B 側の
    UNIT内結線図フィルタ機能専用）。Total シートが無い、または
    『ラベル』『個数』『図番』列が無い場合は ValueError。
    """
    df = _parse_total_sheet(file_bytes, REQUIRED_TOTAL_COLUMNS_WITH_GZUBAN)
    rows = []
    for lbl, cnt, gzuban in zip(df['ラベル'], df['個数'], df['図番']):
        if pd.isna(lbl):
            continue
        key = normalize_label(str(lbl))
        if pd.isna(gzuban):
            gzuban_list = []
        else:
            gzuban_list = [
                normalize_label(g.strip()) for g in str(gzuban).split(',') if g.strip()
            ]
        rows.append((key, int(cnt), gzuban_list))
    return rows


def load_summary_titles(file_bytes: bytes) -> dict:
    """Excel の Summary シートを読み、正規化図番 → タイトル の dict を返す。

    Summary シートが無い、または『図番』『タイトル』列が無い場合は ValueError。
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if REQUIRED_SUMMARY_SHEET not in xls.sheet_names:
        raise ValueError(f"'{REQUIRED_SUMMARY_SHEET}' シートが見つかりません")
    df = xls.parse(REQUIRED_SUMMARY_SHEET)
    missing = [c for c in REQUIRED_SUMMARY_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"{REQUIRED_SUMMARY_SHEET} シートに必要な列がありません: {', '.join(missing)}")

    title_map: dict = {}
    for gzuban, title in zip(df['図番'], df['タイトル']):
        if pd.isna(gzuban):
            continue
        key = normalize_label(str(gzuban).strip())
        title_map[key] = '' if pd.isna(title) else str(title)
    return title_map
