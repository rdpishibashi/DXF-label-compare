import streamlit as st
import pandas as pd

from utils.excel_input import load_total_labels, load_total_rows, load_summary_titles
from utils.compare_labels import (
    compare_labels, summarize, KUBUN_BOTH, KUBUN_A_ONLY, KUBUN_B_ONLY,
)
from utils.excel_output import create_compare_excel_output
from utils.drawing_filter import (
    select_drawing_numbers, aggregate_filtered_rows,
    FILTER_UNIT_ONLY, FILTER_OPTIONS,
)

st.set_page_config(
    page_title="DXF Label Compare",
    page_icon="🔍",
    layout="wide",
)

_BG = {KUBUN_BOTH: '#C6EFCE', KUBUN_A_ONLY: '#D9E1F2', KUBUN_B_ONLY: '#E2CFC0'}
_FG = {KUBUN_BOTH: '#006100', KUBUN_A_ONLY: '#1F4E79', KUBUN_B_ONLY: '#7F4F24'}


def _row_style(row):
    kubun = row['区分']
    css = f'background-color: {_BG[kubun]}; color: {_FG[kubun]}'
    return [css] * len(row)


def _for_display(diff_df):
    """画面表示用に A個数/B個数 を文字列化する（欠損は空文字）。

    st.dataframe は Int64 列の pd.NA をグレーの "None" として描画してしまう
    （Streamlit 1.54 のデータグリッドの仕様。Styler.format(na_rep=...) や
    column_config.NumberColumn でも抑制できない）。表示専用のコピーで
    文字列化し、欠損を空文字にすることで空欄として見せる。
    Excel 出力側（utils/excel_output.py）は write_blank で別途正しく空欄化
    しているため、この処理はダウンロードファイルには影響しない。"""
    disp = diff_df.copy()
    for col in ('A個数', 'B個数'):
        disp[col] = disp[col].apply(lambda v: '' if pd.isna(v) else str(int(v)))
    return disp


def main():
    st.title("DXF Label Compare")
    st.caption(
        "DXF-extract-labels が出力した2つのExcel（Totalシート）のラベルを比較し、"
        "差分（Aのみ・Bのみ・両方）をExcelで出力します。"
    )

    col_a, col_b = st.columns(2)
    with col_a:
        file_a = st.file_uploader(
            "A: 展開接続図", type=['xlsx'], key='uploader_a',
        )
    with col_b:
        file_b = st.file_uploader(
            "B: UNIT内結線図", type=['xlsx'], key='uploader_b',
        )
        filter_mode = FILTER_UNIT_ONLY
        if file_b is not None:
            filter_mode = st.radio(
                "対象範囲（Bのタイトルで絞り込み）",
                FILTER_OPTIONS,
                index=0,
                key='b_filter_mode',
                horizontal=True,
            )

    has_input = file_a is not None and file_b is not None
    run = st.button("比較", type="primary", disabled=not has_input)

    if run:
        try:
            labels_a = load_total_labels(file_a.getvalue())

            b_bytes = file_b.getvalue()
            title_map = load_summary_titles(b_bytes)
            total_rows_b = load_total_rows(b_bytes)
            selected_gzuban = select_drawing_numbers(title_map, filter_mode)
            labels_b = aggregate_filtered_rows(total_rows_b, selected_gzuban)
        except ValueError as e:
            st.error(f"読み込みエラー: {e}")
            return

        diff_df = compare_labels(labels_a, labels_b)
        summary = summarize(diff_df, file_a.name, file_b.name, b_filter_mode=filter_mode)
        excel_bytes = create_compare_excel_output(diff_df, summary)

        st.session_state['diff_df'] = diff_df
        st.session_state['summary'] = summary
        st.session_state['excel_result'] = excel_bytes
        st.session_state['download_done'] = False
        st.rerun()

    if 'diff_df' in st.session_state:
        diff_df = st.session_state['diff_df']
        summary = st.session_state['summary']

        st.divider()
        st.subheader("結果")
        if 'B 絞り込み条件' in summary:
            st.caption(f"B の対象範囲: {summary['B 絞り込み条件']}")
        st.info(
            f"A のみ: {summary['A のみ']}件　/　"
            f"B のみ: {summary['B のみ']}件　/　"
            f"両方: {summary['両方']}件　/　"
            f"ユニーク合計: {summary['ユニーク合計']}件"
        )

        styled = _for_display(diff_df).style.apply(_row_style, axis=1)
        st.dataframe(styled, width='stretch', hide_index=True)

        download_done = st.session_state.get('download_done', False)
        downloaded = st.download_button(
            label="Excelをダウンロード",
            data=st.session_state['excel_result'],
            file_name="label_compare.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            width='stretch',
        )
        if downloaded and not download_done:
            st.session_state['download_done'] = True
            st.rerun()


if __name__ == '__main__':
    main()
