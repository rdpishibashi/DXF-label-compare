"""ラベル比較のコアロジック（Streamlit非依存の純関数）。"""
import pandas as pd

DIFF_COLUMNS = ['ラベル', '区分', 'A個数', 'B個数']
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


def summarize(df: pd.DataFrame, a_name: str, b_name: str) -> dict:
    """サマリー用の件数集計を返す。"""
    a_only = int((df['区分'] == KUBUN_A_ONLY).sum())
    b_only = int((df['区分'] == KUBUN_B_ONLY).sum())
    both = int((df['区分'] == KUBUN_BOTH).sum())
    return {
        'A ファイル名': a_name,
        'B ファイル名': b_name,
        'A ユニークラベル数': both + a_only,
        'B ユニークラベル数': both + b_only,
        'A のみ': a_only,
        'B のみ': b_only,
        '両方': both,
        'ユニーク合計': len(df),
    }
