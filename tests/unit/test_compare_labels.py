import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from model.compare_labels import (
    normalize_label, compare_labels, summarize, summarize_metrics,
    compare_labels_by_region, build_region_summary_rows, REGION_DIFF_COLUMNS,
)


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
    assert 'B 絞り込み条件' not in summary


def test_summarize_includes_b_filter_mode_when_provided():
    a = {'CN1': 2}
    b = {'CN1': 5}
    df = compare_labels(a, b)
    summary = summarize(df, 'A.xlsx', 'B.xlsx', b_filter_mode='UNIT内結線図のみ')

    assert summary['B 絞り込み条件'] == 'UNIT内結線図のみ'
    # 挿入位置: B ファイル名の直後
    keys = list(summary.keys())
    assert keys.index('B 絞り込み条件') == keys.index('B ファイル名') + 1


def test_summarize_metrics_matches_summarize_subset():
    a = {'CN1': 2, 'R10': 1, 'ABC': 3}
    b = {'CN1': 5, 'X1': 1, 'ABC': 4}
    df = compare_labels(a, b)
    metrics = summarize_metrics(df)
    full = summarize(df, 'A.xlsx', 'B.xlsx')
    for key in ('A ユニークラベル数', 'B ユニークラベル数', 'A のみ', 'B のみ', '両方', 'ユニーク合計'):
        assert metrics[key] == full[key]


def test_compare_labels_by_region_single_region():
    a_by_region = {'R1': {'CN1': 2, 'R10': 1}}
    b_by_region = {'R1': {'CN1': 5, 'X1': 1}}
    diff_df, metrics = compare_labels_by_region(a_by_region, b_by_region, ['R1'])

    assert list(diff_df.columns) == REGION_DIFF_COLUMNS
    assert (diff_df['領域名'] == 'R1').all()
    assert sorted(diff_df['ラベル']) == ['CN1', 'R10', 'X1']
    assert metrics == {'R1': {
        'A ユニークラベル数': 2, 'B ユニークラベル数': 2,
        'A のみ': 1, 'B のみ': 1, '両方': 1, 'ユニーク合計': 3,
    }}


def test_compare_labels_by_region_multiple_regions_kept_separate():
    a_by_region = {'R1': {'CN1': 1}, 'R2': {'CN1': 9}}
    b_by_region = {'R1': {'CN1': 1}, 'R2': {}}
    diff_df, metrics = compare_labels_by_region(a_by_region, b_by_region, ['R1', 'R2'])

    # 同じラベルが2つの領域それぞれで独立した行として出力される
    cn1_rows = diff_df[diff_df['ラベル'] == 'CN1']
    assert len(cn1_rows) == 2
    assert set(cn1_rows['領域名']) == {'R1', 'R2'}
    r1_row = cn1_rows[cn1_rows['領域名'] == 'R1'].iloc[0]
    assert r1_row['区分'] == '両方'
    r2_row = cn1_rows[cn1_rows['領域名'] == 'R2'].iloc[0]
    assert r2_row['区分'] == 'A のみ'
    assert metrics['R1']['両方'] == 1
    assert metrics['R2']['A のみ'] == 1
    # regions の順（挿入順）で連結されること
    assert list(diff_df['領域名'].unique()) == ['R1', 'R2']


def test_compare_labels_by_region_missing_region_treated_as_empty():
    # b_by_region に無い領域は空dict扱い（全ラベルが A のみになる）
    a_by_region = {'R1': {'CN1': 1}}
    b_by_region = {}
    diff_df, metrics = compare_labels_by_region(a_by_region, b_by_region, ['R1'])
    assert (diff_df['区分'] == 'A のみ').all()
    assert metrics['R1']['B のみ'] == 0


def test_compare_labels_by_region_empty_regions_list():
    diff_df, metrics = compare_labels_by_region({}, {}, [])
    assert list(diff_df.columns) == REGION_DIFF_COLUMNS
    assert len(diff_df) == 0
    assert metrics == {}


def test_build_region_summary_rows_layout():
    metrics_by_region = {
        'R1': {'A ユニークラベル数': 2, 'B ユニークラベル数': 2, 'A のみ': 1, 'B のみ': 1, '両方': 1, 'ユニーク合計': 3},
        'R2': {'A ユニークラベル数': 5, 'B ユニークラベル数': 0, 'A のみ': 5, 'B のみ': 0, '両方': 0, 'ユニーク合計': 5},
    }
    rows = build_region_summary_rows(metrics_by_region, 'A.xlsx', 'B.xlsx', b_filter_mode='全部')

    # ヘッダー行（領域名は空欄）
    assert rows[0] == {'項目': 'A ファイル名', '領域名': '', '値': 'A.xlsx'}
    assert rows[1] == {'項目': 'B ファイル名', '領域名': '', '値': 'B.xlsx'}
    assert rows[2] == {'項目': 'B 絞り込み条件', '領域名': '', '値': '全部'}
    # 領域ごとに6項目ずつ、metrics_by_region の順で並ぶ
    r1_rows = [r for r in rows if r['領域名'] == 'R1']
    r2_rows = [r for r in rows if r['領域名'] == 'R2']
    assert len(r1_rows) == 6
    assert len(r2_rows) == 6
    assert rows.index(r1_rows[0]) < rows.index(r2_rows[0])
    assert {r['項目'] for r in r1_rows} == {
        'A ユニークラベル数', 'B ユニークラベル数', 'A のみ', 'B のみ', '両方', 'ユニーク合計',
    }


def test_build_region_summary_rows_without_filter_mode_omits_row():
    rows = build_region_summary_rows({'R1': summarize_metrics(compare_labels({}, {}))}, 'A.xlsx', 'B.xlsx')
    assert 'B 絞り込み条件' not in {r['項目'] for r in rows if r['領域名'] == ''}
