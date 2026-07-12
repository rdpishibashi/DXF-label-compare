"""DXF-extract-labels 出力Excelの Total シート読み込み。"""
import io

import pandas as pd

from utils.compare_labels import normalize_label

REQUIRED_SHEET = 'Total'
REQUIRED_COLUMNS = ('ラベル', '個数')


def load_total_labels(file_bytes: bytes) -> dict:
    """Excel の Total シートを読み、正規化ラベル → 合計個数 の dict を返す。

    Total シートが無い、または『ラベル』『個数』列が無い場合は ValueError。
    正規化（全角→半角）で複数の元ラベルが同一キーになった場合は個数を合算する。
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if REQUIRED_SHEET not in xls.sheet_names:
        raise ValueError(f"'{REQUIRED_SHEET}' シートが見つかりません")
    df = xls.parse(REQUIRED_SHEET)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Total シートに必要な列がありません: {', '.join(missing)}")

    agg: dict = {}
    for lbl, cnt in zip(df['ラベル'], df['個数']):
        if pd.isna(lbl):
            continue
        key = normalize_label(str(lbl))
        agg[key] = agg.get(key, 0) + (int(cnt) if pd.notna(cnt) else 0)
    return agg
