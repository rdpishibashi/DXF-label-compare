"""同一Excel内比較の結果Excel出力。

model/excel_output.py（展開図-結線図比較）とシート構成・書式を揃えている
（サマリーシート＋差分シートの2枚、区分ごとの行色分け）。両ファイルの見た目が
再び乖離しないよう、ヘッダー書式・区分の配色・freeze_panes/autofilterの方針を
変更する際は両方のファイルを同時に見直すこと。
"""
import io

import pandas as pd

from model.compare_labels import row_style, ROW_STYLE_COLORS

DIFF_COLUMNS = ["ラベル", "A 合計出現数", "A 図番数", "B 出現数", "A 図番", "比較結果"]


def create_same_workbook_output(result: dict) -> bytes:
    """サマリーシートと差分シートを持つ Excel ファイルを bytes で返す。

    シート順: サマリー → 差分。
    差分シートは表示スタイル区分（青=Aのみ／緑=Bのみ／黄=両方だが個数不一致／
    無色=両方かつ個数一致）ごとに行全体を色分けする（`model.compare_labels.row_style()`）。
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1,
        })
        row_formats = {
            style_key: workbook.add_format({**(color or {}), 'border': 1})
            for style_key, color in ROW_STYLE_COLORS.items()
        }

        _write_summary_sheet(writer, result['summary'], header_fmt)
        _write_diff_sheet(writer, workbook, result['comparison'], header_fmt, row_formats)

    return output.getvalue()


def _write_summary_sheet(writer, summary: dict, header_fmt):
    rows = [{'項目': k, '値': v} for k, v in summary.items()]
    df = pd.DataFrame(rows, columns=['項目', '値'])
    df.to_excel(writer, sheet_name='サマリー', index=False)
    ws = writer.sheets['サマリー']
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name, header_fmt)
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 30)
    ws.freeze_panes(1, 0)


def _write_diff_sheet(writer, workbook, comparison_df: pd.DataFrame, header_fmt, row_formats):
    sheet_name = '差分'
    workbook.add_worksheet(sheet_name)
    ws = writer.sheets[sheet_name]

    for col_idx, col_name in enumerate(DIFF_COLUMNS):
        ws.write(0, col_idx, col_name, header_fmt)

    for row_idx, row in enumerate(comparison_df.itertuples(index=False), start=1):
        label, a_count, a_drawing_count, b_count, a_drawings, kubun = row
        fmt = row_formats[row_style(kubun, a_count, b_count)]
        ws.write(row_idx, 0, label, fmt)
        ws.write_number(row_idx, 1, int(a_count), fmt)
        ws.write_number(row_idx, 2, int(a_drawing_count), fmt)
        ws.write_number(row_idx, 3, int(b_count), fmt)
        ws.write(row_idx, 4, a_drawings, fmt)
        ws.write(row_idx, 5, kubun, fmt)

    ws.set_column('A:A', 30)
    ws.set_column('B:D', 12)
    ws.set_column('E:E', 40)
    ws.set_column('F:F', 12)
    ws.freeze_panes(1, 0)
    last_row = len(comparison_df)
    if last_row > 0:
        ws.autofilter(0, 0, last_row, len(DIFF_COLUMNS) - 1)
