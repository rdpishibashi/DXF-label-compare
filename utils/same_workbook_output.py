"""同一Excel内比較の結果Excel出力。"""
import io

import pandas as pd


_STATUS_COLORS = {
    "共通": {"bg_color": "#E2F0D9", "font_color": "#375623"},
    "Aのみ": {"bg_color": "#FFF2CC", "font_color": "#7F6000"},
    "Bのみ": {"bg_color": "#FCE4D6", "font_color": "#C00000"},
}


def create_same_workbook_output(result: dict) -> bytes:
    """比較結果、A/Bのラベル一覧、Aの対象図番を持つExcelを返す。"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        title_fmt = workbook.add_format({
            "bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "font_size": 14,
        })
        subtitle_fmt = workbook.add_format({
            "italic": True, "font_color": "#1F1F1F", "bg_color": "#D9EAF7", "text_wrap": True,
        })
        header_fmt = workbook.add_format({
            "bold": True, "font_color": "#FFFFFF", "bg_color": "#5B9BD5", "border": 1,
            "align": "center", "valign": "vcenter",
        })
        cell_fmt = workbook.add_format({"border": 1, "valign": "top"})
        wrap_fmt = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
        number_fmt = workbook.add_format({"border": 1, "num_format": "#,##0", "valign": "top"})
        status_fmts = {
            status: workbook.add_format({"border": 1, "valign": "top", **colors})
            for status, colors in _STATUS_COLORS.items()
        }

        _write_comparison(writer, result, title_fmt, subtitle_fmt, header_fmt, cell_fmt, wrap_fmt,
                          number_fmt, status_fmts)
        _write_table_sheet(writer, "A_ラベル一覧", "A：UNIT内結線図の統合・重複除外済みラベル一覧",
                           pd.DataFrame(result["a_labels"]), title_fmt, header_fmt, cell_fmt,
                           wrap_fmt, number_fmt, (34, 14, 14, 52))
        _write_table_sheet(writer, "B_ラベル一覧", f"B：{result['b_sheet_name']} のラベル一覧",
                           pd.DataFrame(result["b_labels"]), title_fmt, header_fmt, cell_fmt,
                           wrap_fmt, number_fmt, (38, 14))
        _write_table_sheet(writer, "A_対象図番", "A：選択対象（タイトルに UNIT内結線図 を含む図番）",
                           pd.DataFrame(result["selected_drawings"]), title_fmt, header_fmt,
                           cell_fmt, wrap_fmt, number_fmt, (20, 24, 40))
    return output.getvalue()


def _write_comparison(writer, result, title_fmt, subtitle_fmt, header_fmt, cell_fmt, wrap_fmt,
                      number_fmt, status_fmts):
    sheet_name = "比較結果"
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    worksheet.hide_gridlines(2)
    worksheet.merge_range("A1:F1", f"ラベル比較：UNIT内結線図（A） vs {result['b_sheet_name']}（B）", title_fmt)
    worksheet.merge_range(
        "A2:F2",
        "A は、タイトルが「UNIT内結線図」（全角・半角を同一視）を含む図番の全ラベルを統合し、重複を除外して並べ替えた一覧です。",
        subtitle_fmt,
    )
    worksheet.set_row(1, 30)
    worksheet.write_row(3, 0, ["集計", "件数"], header_fmt)
    for row_idx, (label, value) in enumerate(result["summary"].items(), start=4):
        worksheet.write(row_idx, 0, label, cell_fmt)
        worksheet.write_number(row_idx, 1, value, number_fmt)

    comparison = result["comparison"]
    header_row = 12
    for col_idx, column in enumerate(comparison.columns):
        worksheet.write(header_row, col_idx, column, header_fmt)
    for row_idx, row in enumerate(comparison.itertuples(index=False), start=header_row + 1):
        label, a_count, a_drawings, b_count, drawings, status = row
        worksheet.write(row_idx, 0, label, cell_fmt)
        worksheet.write_number(row_idx, 1, int(a_count), number_fmt)
        worksheet.write_number(row_idx, 2, int(a_drawings), number_fmt)
        worksheet.write_number(row_idx, 3, int(b_count), number_fmt)
        worksheet.write(row_idx, 4, drawings, wrap_fmt)
        worksheet.write(row_idx, 5, status, status_fmts[status])
    worksheet.set_column("A:A", 34)
    worksheet.set_column("B:D", 14)
    worksheet.set_column("E:E", 52)
    worksheet.set_column("F:F", 14)
    worksheet.freeze_panes(header_row + 1, 0)
    if len(comparison):
        worksheet.autofilter(header_row, 0, header_row + len(comparison), len(comparison.columns) - 1)


def _write_table_sheet(writer, sheet_name, title, df, title_fmt, header_fmt, cell_fmt, wrap_fmt,
                       number_fmt, widths):
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    worksheet.hide_gridlines(2)
    last_col = max(0, len(df.columns) - 1)
    worksheet.merge_range(0, 0, 0, last_col, title, title_fmt)
    for col_idx, column in enumerate(df.columns):
        worksheet.write(2, col_idx, column, header_fmt)
    for row_idx, row in enumerate(df.itertuples(index=False), start=3):
        for col_idx, value in enumerate(row):
            fmt = wrap_fmt if df.columns[col_idx] in ("図番", "サブタイトル") else number_fmt if isinstance(value, int) else cell_fmt
            if isinstance(value, int):
                worksheet.write_number(row_idx, col_idx, value, fmt)
            else:
                worksheet.write(row_idx, col_idx, value, fmt)
    for col_idx, width in enumerate(widths):
        worksheet.set_column(col_idx, col_idx, width)
    worksheet.freeze_panes(3, 0)
    if len(df):
        worksheet.autofilter(2, 0, 2 + len(df), last_col)
