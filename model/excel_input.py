"""DXF-extract-labels 出力Excelの Total/Summary シート読み込み。"""
import io

import pandas as pd

from model.compare_labels import normalize_label

REQUIRED_TOTAL_SHEET = 'Total'
REQUIRED_TOTAL_COLUMNS = ('ラベル', '個数')
REQUIRED_TOTAL_COLUMNS_WITH_GZUBAN = ('ラベル', '個数', '図番')
REQUIRED_SUMMARY_SHEET = 'Summary'
REQUIRED_SUMMARY_COLUMNS = ('図番', 'タイトル')
REGION_SHEET = '領域別ラベル一覧'
REGION_SHEET_FIXED_COLUMNS = ('領域名', 'ラベル', '合計個数')


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
    `model.drawing_filter` の図番フィルタと組み合わせて使う（現状は B 側の
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


def has_region_sheet(file_bytes: bytes) -> bool:
    """Excel が『領域別ラベル一覧』シートを持つか判定する（シート名一覧のみ読む軽量チェック）。"""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return REGION_SHEET in xls.sheet_names


def load_region_rows(file_bytes: bytes) -> list:
    """Excel の『領域別ラベル一覧』シートを読み、
    (正規化領域名, 正規化ラベル, 合計個数, 図番リスト) のタプルのリストで返す。

    シート構成は 領域名/ラベル/合計個数/(図番,個数)×ファイル数
    （`DXF-extract-labels` の `build_region_label_summary` 出力。列名『図番』
    『個数』はファイル数分繰り返されるため pandas が `図番.1`/`個数.1` 等に
    自動リネームする。ここでは列名ではなく位置（3列目以降を2列ずつ）で読む）。
    図番リストには、その行で個数が1以上だった列の図番のみを含める
    （`load_total_rows` の図番リストと同じ「実際に出現したファイルのみ」という
    考え方。個数0の列＝その領域にそのラベルが出現しなかったファイル）。

    シートが無い、または先頭3列が『領域名』『ラベル』『合計個数』でない場合は
    ValueError。
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if REGION_SHEET not in xls.sheet_names:
        raise ValueError(f"'{REGION_SHEET}' シートが見つかりません")
    df = xls.parse(REGION_SHEET)
    if len(df.columns) < 3 or tuple(df.columns[:3]) != REGION_SHEET_FIXED_COLUMNS:
        raise ValueError(f"{REGION_SHEET} シートの列構成が想定と異なります")

    rows = []
    for values in df.itertuples(index=False):
        region_name, label = values[0], values[1]
        if pd.isna(region_name) or pd.isna(label):
            continue
        region_key = normalize_label(str(region_name))
        label_key = normalize_label(str(label))
        total_count = int(values[2]) if pd.notna(values[2]) else 0
        gzuban_list = []
        for gz_idx in range(3, len(values), 2):
            cnt_idx = gz_idx + 1
            if cnt_idx >= len(values):
                break
            gz, cnt = values[gz_idx], values[cnt_idx]
            if pd.isna(gz) or pd.isna(cnt) or int(cnt) <= 0:
                continue
            gzuban_list.append(normalize_label(str(gz).strip()))
        rows.append((region_key, label_key, total_count, gzuban_list))
    return rows
