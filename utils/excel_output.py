"""差分結果の Excel 出力モジュール。"""
import io

import pandas as pd

from utils.compare_labels import DIFF_COLUMNS, KUBUN_A_ONLY, KUBUN_B_ONLY, KUBUN_BOTH

_ROW_COLORS = {
    KUBUN_BOTH: {'bg_color': '#C6EFCE', 'font_color': '#006100'},
    KUBUN_A_ONLY: {'bg_color': '#D9E1F2', 'font_color': '#1F4E79'},
    KUBUN_B_ONLY: {'bg_color': '#E2CFC0', 'font_color': '#7F4F24'},
}


def create_compare_excel_output(diff_df: pd.DataFrame, summary: dict) -> bytes:
    """サマリーシートと差分シートを持つ Excel ファイルを bytes で返す。

    シート順: サマリー → 差分。
    差分シートは区分（両方/A のみ/B のみ）ごとに行全体を色分けする。
    A個数・B個数 が pd.NA の場合は空欄セルとして書き込む（0 とは区別する）。
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1,
        })
        row_formats = {
            kubun: workbook.add_format({**style, 'border': 1})
            for kubun, style in _ROW_COLORS.items()
        }

        _write_summary_sheet(writer, workbook, summary, header_fmt)
        _write_diff_sheet(writer, workbook, diff_df, header_fmt, row_formats)

    return output.getvalue()


def _write_summary_sheet(writer, workbook, summary: dict, header_fmt):
    rows = [{'項目': k, '値': v} for k, v in summary.items()]
    df = pd.DataFrame(rows, columns=['項目', '値'])
    df.to_excel(writer, sheet_name='サマリー', index=False)
    ws = writer.sheets['サマリー']
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name, header_fmt)
    ws.set_column('A:A', 22)
    ws.set_column('B:B', 30)
    ws.freeze_panes(1, 0)


def _write_diff_sheet(writer, workbook, diff_df: pd.DataFrame, header_fmt, row_formats):
    sheet_name = '差分'
    # to_excel で書かせず、セル単位で書式付き書き込みを行うため空シートを作ってから埋める
    workbook.add_worksheet(sheet_name)
    ws = writer.sheets[sheet_name]

    for col_idx, col_name in enumerate(DIFF_COLUMNS):
        ws.write(0, col_idx, col_name, header_fmt)

    for row_idx, row in enumerate(diff_df.itertuples(index=False), start=1):
        label, kubun, a_cnt, b_cnt = row
        fmt = row_formats[kubun]
        ws.write(row_idx, 0, label, fmt)
        ws.write(row_idx, 1, kubun, fmt)
        if pd.notna(a_cnt):
            ws.write_number(row_idx, 2, int(a_cnt), fmt)
        else:
            ws.write_blank(row_idx, 2, None, fmt)
        if pd.notna(b_cnt):
            ws.write_number(row_idx, 3, int(b_cnt), fmt)
        else:
            ws.write_blank(row_idx, 3, None, fmt)

    ws.set_column('A:A', 30)
    ws.set_column('B:B', 12)
    ws.set_column('C:D', 10)
    ws.freeze_panes(1, 0)
    last_row = len(diff_df)
    if last_row > 0:
        ws.autofilter(0, 0, last_row, len(DIFF_COLUMNS) - 1)
