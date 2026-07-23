import io
import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from model.excel_input import has_region_sheet, load_region_rows


def _region_workbook_bytes(rows, file_idents=('EE001', 'EE002')):
    """『領域別ラベル一覧』シート（領域名/ラベル/合計個数/(図番,個数)×N）を持つ
    Excel bytes を組み立てる。rows は
    (領域名, ラベル, 合計個数, {図番: 個数}) のタプル列。
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        header = ['領域名', 'ラベル', '合計個数']
        for _ in file_idents:
            header += ['図番', '個数']
        records = []
        for region, label, total, per_file in rows:
            record = {'領域名': region, 'ラベル': label, '合計個数': total}
            # 重複列名は pandas 側で自動リネームされるため、位置ベースで組み立てる
            values = [region, label, total]
            for ident in file_idents:
                values += [ident, per_file.get(ident, 0)]
            records.append(values)
        df = pd.DataFrame(records, columns=[f'c{i}' for i in range(len(header))])
        df.columns = header  # 重複名のまま書き込む（実際の出力と同じ）
        df.to_excel(writer, sheet_name='領域別ラベル一覧', index=False)
    return output.getvalue()


def _no_region_sheet_workbook_bytes():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pd.DataFrame([{'ラベル': 'A', '個数': 1}]).to_excel(writer, sheet_name='Total', index=False)
    return output.getvalue()


def test_has_region_sheet_true_when_present():
    wb = _region_workbook_bytes([('R1', 'A', 1, {'EE001': 1, 'EE002': 0})])
    assert has_region_sheet(wb) is True


def test_has_region_sheet_false_when_absent():
    assert has_region_sheet(_no_region_sheet_workbook_bytes()) is False


def test_load_region_rows_basic():
    wb = _region_workbook_bytes([
        ('R1', 'A', 3, {'EE001': 2, 'EE002': 1}),
        ('R1', 'B', 1, {'EE001': 0, 'EE002': 1}),
        ('R2', 'A', 5, {'EE001': 5, 'EE002': 0}),
    ])
    rows = load_region_rows(wb)
    assert rows == [
        ('R1', 'A', 3, ['EE001', 'EE002']),
        ('R1', 'B', 1, ['EE002']),
        ('R2', 'A', 5, ['EE001']),
    ]


def test_load_region_rows_normalizes_region_and_label():
    wb = _region_workbook_bytes([('Ｒ１', 'Ａ１', 2, {'EE001': 2})])
    rows = load_region_rows(wb)
    assert rows == [('R1', 'A1', 2, ['EE001'])]


def test_load_region_rows_missing_sheet_raises():
    with pytest.raises(ValueError):
        load_region_rows(_no_region_sheet_workbook_bytes())
