import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from model.drawing_filter import (
    select_drawing_numbers, aggregate_filtered_rows,
    FILTER_UNIT_ONLY, FILTER_UNIT_EXCLUDED, FILTER_ALL,
)


TITLE_MAP = {
    'EE1': 'ＵＮＩＴ内結線図',   # 全角UNIT
    'EE2': 'UNIT内結線図',       # 半角UNIT
    'EE3': '展開接続図',         # 非UNIT
    'EE4': '',                   # タイトル空欄
}


def test_select_unit_only_matches_both_fullwidth_and_halfwidth():
    result = select_drawing_numbers(TITLE_MAP, FILTER_UNIT_ONLY)
    assert result == {'EE1', 'EE2'}


def test_select_unit_excluded_includes_non_unit_and_blank_title():
    result = select_drawing_numbers(TITLE_MAP, FILTER_UNIT_EXCLUDED)
    assert result == {'EE3', 'EE4'}


def test_select_all_returns_none():
    result = select_drawing_numbers(TITLE_MAP, FILTER_ALL)
    assert result is None


def test_select_invalid_mode_raises():
    import pytest
    with pytest.raises(ValueError):
        select_drawing_numbers(TITLE_MAP, '不明なモード')


def test_aggregate_filtered_rows_with_selection():
    rows = [
        ('CN1', 5, ['EE1', 'EE3']),  # UNIT(EE1) と 非UNIT(EE3) の両方に出現
        ('R10', 2, ['EE3']),          # 非UNITのみ
        ('X1', 1, ['EE1']),           # UNITのみ
        ('Y1', 3, []),                 # 図番情報なし
    ]
    selected = {'EE1', 'EE2'}  # UNIT内結線図のみ選択時
    result = aggregate_filtered_rows(rows, selected)
    # CN1: EE1が選択集合に含まれるため対象。R10: EE3のみで対象外。
    # X1: EE1が対象。Y1: 図番リストが空のため対象外。
    assert result == {'CN1': 5, 'X1': 1}


def test_aggregate_filtered_rows_excluded_mode():
    rows = [
        ('CN1', 5, ['EE1', 'EE3']),
        ('R10', 2, ['EE3']),
        ('X1', 1, ['EE1']),
    ]
    selected = {'EE3', 'EE4'}  # UNIT内結線図以外選択時
    result = aggregate_filtered_rows(rows, selected)
    assert result == {'CN1': 5, 'R10': 2}


def test_aggregate_filtered_rows_none_selection_includes_all():
    rows = [
        ('CN1', 5, ['EE1']),
        ('R10', 2, []),
    ]
    result = aggregate_filtered_rows(rows, None)
    assert result == {'CN1': 5, 'R10': 2}


def test_aggregate_filtered_rows_sums_duplicate_labels():
    rows = [
        ('CN1', 2, ['EE1']),
        ('CN1', 3, ['EE2']),
    ]
    selected = {'EE1', 'EE2'}
    result = aggregate_filtered_rows(rows, selected)
    assert result == {'CN1': 5}
