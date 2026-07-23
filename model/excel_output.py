"""差分結果の Excel 出力モジュール。"""
import io

import pandas as pd

from model.compare_labels import DIFF_COLUMNS, REGION_DIFF_COLUMNS, KUBUN_A_ONLY, KUBUN_B_ONLY, KUBUN_BOTH

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
    summary_rows = [{'項目': k, '値': v} for k, v in summary.items()]
    return _create_excel_output(diff_df, summary_rows, DIFF_COLUMNS)


def create_region_compare_excel_output(diff_df: pd.DataFrame, summary_rows: list) -> bytes:
    """指定領域での比較結果の Excel ファイルを bytes で返す。

    `create_compare_excel_output` と同じシート構成・配色だが、差分シートの
    先頭に『領域名』列（`REGION_DIFF_COLUMNS`）を持ち、サマリーシートは
    項目・領域名・値の3列（`build_region_summary_rows()` の戻り値をそのまま渡す）。
    """
    return _create_excel_output(diff_df, summary_rows, REGION_DIFF_COLUMNS)


def _create_excel_output(diff_df: pd.DataFrame, summary_rows: list, diff_columns: list) -> bytes:
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

        _write_summary_sheet(writer, summary_rows, header_fmt)
        _write_diff_sheet(writer, diff_df, diff_columns, header_fmt, row_formats)

    return output.getvalue()


def _write_summary_sheet(writer, summary_rows: list, header_fmt):
    columns = list(summary_rows[0].keys()) if summary_rows else ['項目', '値']
    df = pd.DataFrame(summary_rows, columns=columns)
    df.to_excel(writer, sheet_name='サマリー', index=False)
    ws = writer.sheets['サマリー']
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name, header_fmt)
    ws.set_column(0, 0, 22)
    if '領域名' in columns:
        ws.set_column(1, 1, 25)
        ws.set_column(2, 2, 30)
    else:
        ws.set_column(1, 1, 30)
    ws.freeze_panes(1, 0)


def _write_diff_sheet(writer, diff_df: pd.DataFrame, columns: list, header_fmt, row_formats):
    sheet_name = '差分'
    # to_excel で書かせず、セル単位で書式付き書き込みを行うため空シートを作ってから埋める
    workbook = writer.book
    workbook.add_worksheet(sheet_name)
    ws = writer.sheets[sheet_name]

    for col_idx, col_name in enumerate(columns):
        ws.write(0, col_idx, col_name, header_fmt)

    kubun_idx = columns.index('区分')
    a_idx = columns.index('A個数')
    b_idx = columns.index('B個数')

    for row_idx, row in enumerate(diff_df.itertuples(index=False), start=1):
        values = list(row)
        fmt = row_formats[values[kubun_idx]]
        for col_idx, value in enumerate(values):
            if col_idx in (a_idx, b_idx):
                if pd.notna(value):
                    ws.write_number(row_idx, col_idx, int(value), fmt)
                else:
                    ws.write_blank(row_idx, col_idx, None, fmt)
            else:
                ws.write(row_idx, col_idx, value, fmt)

    if '領域名' in columns:
        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 30)
        ws.set_column(2, 2, 12)
        ws.set_column(3, 4, 10)
    else:
        ws.set_column(0, 0, 30)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 3, 10)
    ws.freeze_panes(1, 0)
    last_row = len(diff_df)
    if last_row > 0:
        ws.autofilter(0, 0, last_row, len(columns) - 1)
