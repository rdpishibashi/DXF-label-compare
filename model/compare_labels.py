"""ラベル比較のコアロジック（Streamlit非依存の純関数）。"""
import pandas as pd

DIFF_COLUMNS = ['ラベル', '区分', 'A個数', 'B個数']
REGION_DIFF_COLUMNS = ['領域名'] + DIFF_COLUMNS
KUBUN_BOTH = '両方'
KUBUN_A_ONLY = 'A のみ'
KUBUN_B_ONLY = 'B のみ'


def normalize_label(s: str) -> str:
    """全角ASCII(U+FF01-FF5E)を半角に、全角スペースを半角スペースに変換する。
    日本語文字（かな・カナ・漢字）は変換しない。"""
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            out.append(chr(o - 0xFEE0))
        elif o == 0x3000:
            out.append(' ')
        else:
            out.append(ch)
    return ''.join(out)


def compare_labels(a: dict, b: dict) -> pd.DataFrame:
    """正規化済みの {ラベル: 個数} 2 つを比較し、差分 DataFrame を返す。

    - columns: ['ラベル', '区分', 'A個数', 'B個数']（この順）
    - 区分 ∈ {'A のみ', 'B のみ', '両方'}
    - 無い側の個数は pd.NA
    - ラベル昇順（sorted）
    """
    labels = sorted(set(a) | set(b))
    rows = []
    for lbl in labels:
        in_a, in_b = lbl in a, lbl in b
        kubun = KUBUN_BOTH if (in_a and in_b) else (KUBUN_A_ONLY if in_a else KUBUN_B_ONLY)
        rows.append({
            'ラベル': lbl,
            '区分': kubun,
            'A個数': a.get(lbl, pd.NA),
            'B個数': b.get(lbl, pd.NA),
        })
    df = pd.DataFrame(rows, columns=DIFF_COLUMNS)
    # 個数は pandas の nullable 整数型にする。素の object dtype のままだと
    # st.dataframe（Arrow経由の描画）で pd.NA が文字列 "None" として
    # 表示されてしまうため（Excel出力側は write_blank で別途空欄化している）。
    df['A個数'] = df['A個数'].astype('Int64')
    df['B個数'] = df['B個数'].astype('Int64')
    return df


def summarize_metrics(df: pd.DataFrame) -> dict:
    """区分カウントのみを返す（ファイル名等のヘッダー情報を含まない軽量版）。

    'A ユニークラベル数'〜'ユニーク合計' の6項目。`summarize()` と
    `build_region_summary_rows()`（指定領域での比較、領域ごとの内訳）が共用する。
    """
    a_only = int((df['区分'] == KUBUN_A_ONLY).sum())
    b_only = int((df['区分'] == KUBUN_B_ONLY).sum())
    both = int((df['区分'] == KUBUN_BOTH).sum())
    return {
        'A ユニークラベル数': both + a_only,
        'B ユニークラベル数': both + b_only,
        'A のみ': a_only,
        'B のみ': b_only,
        '両方': both,
        'ユニーク合計': len(df),
    }


def summarize(df: pd.DataFrame, a_name: str, b_name: str, b_filter_mode: str = None) -> dict:
    """サマリー用の件数集計を返す。

    b_filter_mode を指定すると、B側の図番フィルタ条件（'B 絞り込み条件'）を
    'B ファイル名' の直後に記録する（未指定時は省略）。
    """
    result = {
        'A ファイル名': a_name,
        'B ファイル名': b_name,
    }
    if b_filter_mode is not None:
        result['B 絞り込み条件'] = b_filter_mode
    result.update(summarize_metrics(df))
    return result


def compare_labels_by_region(a_by_region: dict, b_by_region: dict, regions: list):
    """指定した領域名（regions、表示順）ごとに `compare_labels()` を実行し、
    先頭に『領域名』列を持つ1つの DataFrame に結合する。

    a_by_region/b_by_region: {領域名: {ラベル: 個数}}（該当領域が無ければ空dict扱い）。
    戻り値: (diff_df, metrics_by_region)
      diff_df: columns=REGION_DIFF_COLUMNS。regions の順で連結（regions内では
               ラベル昇順）。regions が空なら0行のDataFrame。
      metrics_by_region: {領域名: summarize_metrics()の戻り値}（regions と同じ順）
    """
    frames = []
    metrics_by_region: dict = {}
    for region in regions:
        a_labels = a_by_region.get(region, {})
        b_labels = b_by_region.get(region, {})
        region_df = compare_labels(a_labels, b_labels)
        metrics_by_region[region] = summarize_metrics(region_df)
        region_df = region_df.copy()
        region_df.insert(0, '領域名', region)
        frames.append(region_df)
    if frames:
        diff_df = pd.concat(frames, ignore_index=True)
    else:
        diff_df = pd.DataFrame(columns=REGION_DIFF_COLUMNS)
    return diff_df, metrics_by_region


def build_region_summary_rows(
    metrics_by_region: dict, a_name: str, b_name: str, b_filter_mode: str = None,
) -> list:
    """領域名ごとの `summarize_metrics()` 結果から、サマリーシート用の行リスト
    （項目・領域名・値の3キーを持つ dict のリスト）を構築する。

    metrics_by_region は表示したい順（`compare_labels_by_region()` の戻り値の
    2番目の要素）。先頭にファイル名等のヘッダー行（領域名は空欄）を置き、
    続けて領域名ごとに6項目（A ユニークラベル数〜ユニーク合計）のブロックを
    metrics_by_region の順で並べる。
    """
    rows = [
        {'項目': 'A ファイル名', '領域名': '', '値': a_name},
        {'項目': 'B ファイル名', '領域名': '', '値': b_name},
    ]
    if b_filter_mode is not None:
        rows.append({'項目': 'B 絞り込み条件', '領域名': '', '値': b_filter_mode})
    for region, metrics in metrics_by_region.items():
        for key, value in metrics.items():
            rows.append({'項目': key, '領域名': region, '値': value})
    return rows
