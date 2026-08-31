import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from model.compare_labels import (
    normalize_label, compare_labels, summarize, summarize_metrics,
    compare_labels_by_region, build_region_summary_rows, REGION_DIFF_COLUMNS,
    row_style, ROW_STYLE_A_ONLY, ROW_STYLE_B_ONLY, ROW_STYLE_MATCH, ROW_STYLE_MISMATCH,
    KUBUN_A_ONLY, KUBUN_B_ONLY, KUBUN_BOTH,
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
    assert 'ユニーク合計' not in summary
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
    for key in ('A ユニークラベル数', 'B ユニークラベル数', 'A のみ', 'B のみ', '両方'):
        assert metrics[key] == full[key]
    assert 'ユニーク合計' not in metrics


def test_compare_labels_by_region_single_region():
    a_by_region = {'R1': {'CN1': 2, 'R10': 1}}
    b_by_region = {'R1': {'CN1': 5, 'X1': 1}}
    diff_df, metrics = compare_labels_by_region(a_by_region, b_by_region, ['R1'])

    assert list(diff_df.columns) == REGION_DIFF_COLUMNS
    assert (diff_df['領域名'] == 'R1').all()
    assert sorted(diff_df['ラベル']) == ['CN1', 'R10', 'X1']
    assert metrics == {'R1': {
        'A ユニークラベル数': 2, 'B ユニークラベル数': 2,
        'A のみ': 1, 'B のみ': 1, '両方': 1,
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
        'R1': {'A ユニークラベル数': 2, 'B ユニークラベル数': 2, 'A のみ': 1, 'B のみ': 1, '両方': 1},
        'R2': {'A ユニークラベル数': 5, 'B ユニークラベル数': 0, 'A のみ': 5, 'B のみ': 0, '両方': 0},
    }
    rows = build_region_summary_rows(metrics_by_region, 'A.xlsx', 'B.xlsx', b_filter_mode='全部')

    # 列順は 領域名・項目・値（領域名が先頭）。ヘッダー行は領域名が空欄
    assert rows[0] == {'領域名': '', '項目': 'A ファイル名', '値': 'A.xlsx'}
    assert rows[1] == {'領域名': '', '項目': 'B ファイル名', '値': 'B.xlsx'}
    assert rows[2] == {'領域名': '', '項目': 'B 絞り込み条件', '値': '全部'}
    # 領域ごとに5項目ずつ、metrics_by_region の順で並ぶ
    r1_block = rows[3:8]
    r2_block = rows[8:13]
    assert len(rows) == 13
    assert {r['項目'] for r in r1_block} == {
        'A ユニークラベル数', 'B ユニークラベル数', 'A のみ', 'B のみ', '両方',
    }
    assert {r['項目'] for r in r2_block} == {
        'A ユニークラベル数', 'B ユニークラベル数', 'A のみ', 'B のみ', '両方',
    }
    # ブロック内は先頭行だけ領域名が入り、残りは空欄
    assert r1_block[0]['領域名'] == 'R1'
    assert all(r['領域名'] == '' for r in r1_block[1:])
    assert r2_block[0]['領域名'] == 'R2'
    assert all(r['領域名'] == '' for r in r2_block[1:])


def test_build_region_summary_rows_without_filter_mode_omits_row():
    rows = build_region_summary_rows({'R1': summarize_metrics(compare_labels({}, {}))}, 'A.xlsx', 'B.xlsx')
    assert 'B 絞り込み条件' not in {r['項目'] for r in rows if r['領域名'] == ''}


# row_style(): 区分 × 個数一致/不一致 の組み合わせ（青=Aのみ／緑=Bのみ／
# 黄=両方だが個数不一致／無色=両方かつ個数一致）。区分が『A のみ』『B のみ』の
# 場合は個数の値に関わらず短絡することも確認する。
def test_row_style_a_only_ignores_counts():
    assert row_style(KUBUN_A_ONLY, 3, pd.NA) == ROW_STYLE_A_ONLY
    assert row_style(KUBUN_A_ONLY, 0, 0) == ROW_STYLE_A_ONLY


def test_row_style_b_only_ignores_counts():
    assert row_style(KUBUN_B_ONLY, pd.NA, 3) == ROW_STYLE_B_ONLY
    assert row_style(KUBUN_B_ONLY, 0, 0) == ROW_STYLE_B_ONLY


def test_row_style_both_matching_counts_is_match():
    assert row_style(KUBUN_BOTH, 5, 5) == ROW_STYLE_MATCH
    assert row_style(KUBUN_BOTH, 0, 0) == ROW_STYLE_MATCH


def test_row_style_both_mismatched_counts_is_mismatch():
    assert row_style(KUBUN_BOTH, 1, 99) == ROW_STYLE_MISMATCH
    assert row_style(KUBUN_BOTH, 99, 1) == ROW_STYLE_MISMATCH


def test_row_style_both_with_missing_count_is_mismatch():
    # 区分が『両方』なのに片方の個数が欠損している状態は本来生じないが、
    # 防御的に不一致（黄）として扱う。
    assert row_style(KUBUN_BOTH, pd.NA, 5) == ROW_STYLE_MISMATCH
    assert row_style(KUBUN_BOTH, 5, pd.NA) == ROW_STYLE_MISMATCH
