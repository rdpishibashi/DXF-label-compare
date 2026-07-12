import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from utils.compare_labels import normalize_label, compare_labels, summarize


def test_basic_kubun_assignment():
    a = {'CN1': 2, 'R10': 1, 'ABC': 3}
    b = {'CN1': 5, 'X1': 1, 'ABC': 4}
    df = compare_labels(a, b)

    assert list(df['ラベル']) == ['ABC', 'CN1', 'R10', 'X1']

    row = {r['ラベル']: r for r in df.to_dict('records')}
    assert row['ABC']['区分'] == '両方'
    assert row['ABC']['A個数'] == 3
    assert row['ABC']['B個数'] == 4

    assert row['CN1']['区分'] == '両方'
    assert row['CN1']['A個数'] == 2
    assert row['CN1']['B個数'] == 5

    assert row['R10']['区分'] == 'A のみ'
    assert row['R10']['A個数'] == 1
    assert pd.isna(row['R10']['B個数'])

    assert row['X1']['区分'] == 'B のみ'
    assert pd.isna(row['X1']['A個数'])
    assert row['X1']['B個数'] == 1


def test_kubun_both_ignores_count_difference():
    a = {'ABC': 1}
    b = {'ABC': 99}
    df = compare_labels(a, b)
    assert len(df) == 1
    assert df.iloc[0]['区分'] == '両方'
    assert df.iloc[0]['A個数'] == 1
    assert df.iloc[0]['B個数'] == 99


def test_normalize_label_fullwidth_ascii_to_halfwidth():
    assert normalize_label('ＣＮ１') == 'CN1'
    assert normalize_label('AB１２') == 'AB12'
    assert normalize_label('あ　い') == 'あ い'
    assert normalize_label('抵抗Ｒ') == '抵抗R'


def test_normalize_label_preserves_japanese():
    assert normalize_label('ラベル') == 'ラベル'
    assert normalize_label('漢字テスト') == '漢字テスト'


def test_normalization_merges_duplicate_keys_and_sums_counts():
    # excel_input.load_total_labels 相当の集約ロジックを直接検証する
    raw = [('ＣＮ１', 2), ('CN1', 3), ('ABC', 1)]
    agg = {}
    for lbl, cnt in raw:
        key = normalize_label(lbl)
        agg[key] = agg.get(key, 0) + cnt
    assert agg == {'CN1': 5, 'ABC': 1}


def test_summarize_counts_are_consistent():
    a = {'CN1': 2, 'R10': 1, 'ABC': 3}
    b = {'CN1': 5, 'X1': 1, 'ABC': 4}
    df = compare_labels(a, b)
    summary = summarize(df, 'A.xlsx', 'B.xlsx')

    assert summary['A ファイル名'] == 'A.xlsx'
    assert summary['B ファイル名'] == 'B.xlsx'
    assert summary['A のみ'] == 1
    assert summary['B のみ'] == 1
    assert summary['両方'] == 2
    assert summary['A ユニークラベル数'] == 3
    assert summary['B ユニークラベル数'] == 3
    assert summary['ユニーク合計'] == 4
    assert summary['ユニーク合計'] == len(df)
