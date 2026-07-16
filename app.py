import streamlit as st
import pandas as pd

from model.excel_input import load_total_labels, load_total_rows, load_summary_titles
from model.compare_labels import (
    compare_labels, summarize, KUBUN_BOTH, KUBUN_A_ONLY, KUBUN_B_ONLY,
)
from model.excel_output import create_compare_excel_output
from model.drawing_filter import (
    select_drawing_numbers, aggregate_filtered_rows,
    FILTER_UNIT_ONLY, FILTER_OPTIONS,
)
from model.same_workbook import compare_within_workbook, list_label_sheets
from model.same_workbook_output import create_same_workbook_output

st.set_page_config(page_title="DXF Label Compare", page_icon="🔍", layout="wide")

_BG = {KUBUN_BOTH: '#C6EFCE', KUBUN_A_ONLY: '#D9E1F2', KUBUN_B_ONLY: '#E2CFC0'}
_FG = {KUBUN_BOTH: '#006100', KUBUN_A_ONLY: '#1F4E79', KUBUN_B_ONLY: '#7F4F24'}


def _row_style(row):
    kubun = row['区分']
    css = f'background-color: {_BG[kubun]}; color: {_FG[kubun]}'
    return [css] * len(row)


def _for_display(diff_df):
    """画面表示用に欠損の個数を空欄にする。"""
    disp = diff_df.copy()
    for col in ('A個数', 'B個数'):
        disp[col] = disp[col].apply(lambda value: '' if pd.isna(value) else str(int(value)))
    return disp


def _show_same_workbook_comparison():
    st.subheader("結線図-組立図比較")
    st.caption(
        "UNIT内結線図ファイルセットから抽出したラベル・ファイルを利用して、"
        "UNIT内結線図に含まれれるラベルと組立図に含まれるラベルを比較します。"
    )
    uploaded = st.file_uploader("抽出ラベルExcel", type=['xlsx'], key='same_workbook_file')
    if uploaded is None:
        return

    workbook_bytes = uploaded.getvalue()
    try:
        label_sheets = list_label_sheets(workbook_bytes)
    except Exception as error:
        st.error(f"Excelを読み込めません: {error}")
        return
    if not label_sheets:
        st.error("個別図面のラベルシート（「ラベル」「個数」列）が見つかりません。")
        return

    b_sheet_name = st.selectbox("B：比較するシート", label_sheets, key='same_workbook_b_sheet')
    if st.button("結線図-組立図比較", type='primary', key='same_workbook_run'):
        try:
            result = compare_within_workbook(workbook_bytes, b_sheet_name)
            st.session_state['same_workbook_result'] = result
            st.session_state['same_workbook_file_name'] = uploaded.name
            st.session_state['same_workbook_output'] = create_same_workbook_output(result)
        except ValueError as error:
            st.error(f"比較できません: {error}")
            return

    result = st.session_state.get('same_workbook_result')
    if result is None:
        return
    st.divider()
    st.subheader("結果")
    summary = result['summary']
    st.info(
        f"A 対象図番: {summary['A 対象図番数']}件　/　"
        f"共通: {summary['共通ラベル数']}件　/　"
        f"A のみ: {summary['A のみ']}件　/　B のみ: {summary['B のみ']}件"
    )
    st.dataframe(result['comparison'], width='stretch', hide_index=True)
    st.download_button(
        label="比較結果Excelをダウンロード",
        data=st.session_state['same_workbook_output'],
        file_name=f"label_comparison_UNIT内結線図_vs_{result['b_sheet_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type='primary',
        width='stretch',
        key='same_workbook_download',
    )


def _show_two_workbook_comparison():
    st.subheader("展開図-結線図比較")
    st.caption(
        "展開接続図ファイルから抽出したラベル・ファイル(A)とUNIT内結線図ファイルセットから"
        "抽出したラベル・ファイル(B)のラベル（機器符号）を比較します。"
    )
    col_a, col_b = st.columns(2)
    with col_a:
        file_a = st.file_uploader("A: 展開接続図のラベル・ファイル", type=['xlsx'], key='uploader_a')
    with col_b:
        file_b = st.file_uploader("B: UNIT内結線図セットのラベル・ファイル", type=['xlsx'], key='uploader_b')
        filter_mode = FILTER_UNIT_ONLY
        if file_b is not None:
            filter_mode = st.radio("対象範囲（Bのタイトルで絞り込み）", FILTER_OPTIONS, index=0,
                                   key='b_filter_mode', horizontal=True)

    if st.button("展開図-結線図比較", type='primary', disabled=file_a is None or file_b is None,
                 key='two_workbook_run'):
        try:
            labels_a = load_total_labels(file_a.getvalue())
            b_bytes = file_b.getvalue()
            title_map = load_summary_titles(b_bytes)
            total_rows_b = load_total_rows(b_bytes)
            selected_gzuban = select_drawing_numbers(title_map, filter_mode)
            labels_b = aggregate_filtered_rows(total_rows_b, selected_gzuban)
            diff_df = compare_labels(labels_a, labels_b)
            summary = summarize(diff_df, file_a.name, file_b.name, b_filter_mode=filter_mode)
            st.session_state['two_workbook_diff'] = diff_df
            st.session_state['two_workbook_summary'] = summary
            st.session_state['two_workbook_output'] = create_compare_excel_output(diff_df, summary)
        except ValueError as error:
            st.error(f"読み込みエラー: {error}")
            return

    diff_df = st.session_state.get('two_workbook_diff')
    if diff_df is None:
        return
    summary = st.session_state['two_workbook_summary']
    st.divider()
    st.subheader("結果")
    st.info(
        f"A のみ: {summary['A のみ']}件　/　B のみ: {summary['B のみ']}件　/　"
        f"両方: {summary['両方']}件　/　ユニーク合計: {summary['ユニーク合計']}件"
    )
    st.dataframe(_for_display(diff_df).style.apply(_row_style, axis=1), width='stretch', hide_index=True)
    st.download_button(
        label="Excelをダウンロード", data=st.session_state['two_workbook_output'],
        file_name="label_compare.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type='primary', width='stretch', key='two_workbook_download',
    )


def main():
    st.markdown("""
        <style>
        @font-face {
            font-family: "AppMixedFont";
            src: local("Hiragino Kaku Gothic ProN"), local("Yu Gothic UI"),
                 local("Yu Gothic"), local("Meiryo");
            unicode-range: U+3000-303F,
                           U+3040-30FF,
                           U+FF00-FFEF,
                           U+4E00-9FFF, U+3400-4DBF;
            size-adjust: 94%;
        }
        @font-face {
            font-family: "AppMixedFont";
            src: local("Source Sans Pro"), local("Helvetica Neue"), local("Arial");
        }
        .stApp, .stApp p, .stApp li, .stApp label, .stApp td, .stApp th,
        .stApp h1, .stApp h2, .stApp h3, .stApp input, .stApp button {
            font-family: "AppMixedFont", sans-serif !important;
        }
        </style>
    """, unsafe_allow_html=True)
    st.title("DXF Label Compare - 機器符号比較")
    st.markdown("""
        <style>
        .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
            align-items: flex-end;
        }
        .stTabs button[data-baseweb="tab"] {
            border: 1px solid rgba(128, 128, 128, 0.5);
            border-bottom: none;
            border-radius: 10px 10px 0 0;
            padding: 4px 20px;
            background: rgba(128, 128, 128, 0.12);
        }
        .stTabs button[data-baseweb="tab"][aria-selected="true"] {
            background: transparent;
        }
        .stTabs button[data-baseweb="tab"] [data-testid="stMarkdownContainer"] p {
            font-size: 16.7px;
        }
        .stTabs button[data-baseweb="tab"][aria-selected="true"] [data-testid="stMarkdownContainer"] p {
            font-weight: 700;
        }
        </style>
    """, unsafe_allow_html=True)
    two_tab, same_tab = st.tabs(["展開図-結線図比較", "結線図-組立図比較"])
    with two_tab:
        _show_two_workbook_comparison()
    with same_tab:
        _show_same_workbook_comparison()


if __name__ == '__main__':
    main()
