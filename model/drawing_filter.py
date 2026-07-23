"""B側（UNIT内結線図）の図番フィルタリング（Streamlit非依存の純関数）。"""
from model.compare_labels import normalize_label

UNIT_TITLE = 'UNIT内結線図'

FILTER_UNIT_ONLY = 'UNIT内結線図のみ'
FILTER_UNIT_EXCLUDED = 'UNIT内結線図以外'
FILTER_ALL = '全部'
FILTER_OPTIONS = (FILTER_UNIT_ONLY, FILTER_UNIT_EXCLUDED, FILTER_ALL)


def select_drawing_numbers(title_map: dict, filter_mode: str):
    """図番→タイトル の対応から、filter_mode に応じた対象図番集合を返す。

    タイトルは正規化（全角→半角）した上で 'UNIT内結線図' と比較するため、
    「ＵＮＩＴ内結線図」（全角）と「UNIT内結線図」（半角）は同じものとして扱う。
    filter_mode == FILTER_ALL の場合は None（絞り込みなし）を返す。
    """
    if filter_mode == FILTER_ALL:
        return None
    unit_gzuban = {
        g for g, title in title_map.items()
        if normalize_label(str(title)) == UNIT_TITLE
    }
    if filter_mode == FILTER_UNIT_ONLY:
        return unit_gzuban
    if filter_mode == FILTER_UNIT_EXCLUDED:
        return set(title_map) - unit_gzuban
    raise ValueError(f"不明な filter_mode: {filter_mode}")


def aggregate_filtered_rows(rows, selected_gzuban):
    """(ラベル, 個数, 図番リスト) のタプル列を集約し、正規化ラベル→合計個数 の dict を返す。

    selected_gzuban が None なら全件を対象にする。selected_gzuban が集合の場合、
    その図番リストが selected_gzuban と1つでも重なる行だけを対象にする
    （図番リストが空、またはどれも selected_gzuban に含まれない行は除外）。
    """
    agg: dict = {}
    for label, count, gzuban_list in rows:
        if selected_gzuban is not None and not (set(gzuban_list) & selected_gzuban):
            continue
        agg[label] = agg.get(label, 0) + count
    return agg


def aggregate_region_rows(rows, selected_gzuban):
    """(領域名, ラベル, 個数, 図番リスト) のタプル列を集約し、
    領域名 -> {ラベル: 合計個数} の2重dictを返す。

    絞り込み方針は `aggregate_filtered_rows` と同じ（selected_gzuban が None なら
    全件対象、集合の場合は図番リストとの重なりで行単位に採否を決め、個数は
    絞り込み後も再計算しない）。領域名は指定領域での比較機能の「共通領域名」の
    候補集合を作る側（呼び出し元）が別途 A/B の領域名集合の積を取るため、
    ここでは行に含まれる領域名をそのまま使う。
    """
    result: dict = {}
    for region, label, count, gzuban_list in rows:
        if selected_gzuban is not None and not (set(gzuban_list) & selected_gzuban):
            continue
        region_dict = result.setdefault(region, {})
        region_dict[label] = region_dict.get(label, 0) + count
    return result
