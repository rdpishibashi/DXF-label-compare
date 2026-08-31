import streamlit as st
import pandas as pd

from model.excel_input import (
    load_total_labels, load_total_rows, load_summary_titles,
    has_region_sheet, load_region_rows,
)
from model.compare_labels import (
    compare_labels, summarize, compare_labels_by_region, build_region_summary_rows,
    blank_repeated_column, row_style, ROW_STYLE_COLORS,
)
from model.excel_output import create_compare_excel_output, create_region_compare_excel_output
from model.drawing_filter import (
    select_drawing_numbers, aggregate_filtered_rows, aggregate_region_rows,
    FILTER_UNIT_ONLY, FILTER_OPTIONS,
)
from model.same_workbook import compare_within_workbook, list_label_sheets
from model.same_workbook_output import create_same_workbook_output

st.set_page_config(page_title="DXF Label Compare", page_icon="🔍", layout="wide")

def _row_style_factory(diff_df):
    """`_for_display()` は A個数/B個数 を表示用文字列に変換するため、色分けの
    判定（`row_style()`、個数の一致比較を含む）は変換前の元データを使う。
    `.style.apply` に渡す行（表示用DataFrameの行）とはインデックスで対応させる。
    """
    def _row_style(row):
        raw = diff_df.loc[row.name]
        style_key = row_style(raw['区分'], raw['A個数'], raw['B個数'])
        colors = ROW_STYLE_COLORS[style_key]
        css = f"background-color: {colors['bg_color']}; color: {colors['font_color']}" if colors else ''
        return [css] * len(row)
    return _row_style


def _for_display(diff_df):
    """画面表示用に欠損の個数を空欄にし、『領域名』列があれば連続する重複を空欄にする。"""
    disp = diff_df.copy()
    for col in ('A個数', 'B個数'):
        disp[col] = disp[col].apply(lambda value: '' if pd.isna(value) else str(int(value)))
    if '領域名' in disp.columns:
        disp = blank_repeated_column(disp, '領域名')
    return disp


def _region_summary_for_display(summary):
    """画面表示用: `build_region_summary_rows()` の行データを DataFrame 化する。

    『値』列はファイル名等の文字列と集計件数の整数が混在する object dtype に
    なるため、st.dataframe（Arrow経由の描画）に渡すと
    `pyarrow.lib.ArrowTypeError` が発生する（コンソールに出るだけで自動
    フォールバックにより見た目上は動作するが、コンソールを汚す）。表示直前に
    文字列化して回避する。Excel出力用の `summary`（生データ）はそのまま変更しない。
    """
    disp = pd.DataFrame(summary)
    disp['値'] = disp['値'].astype(str)
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
        f"両方: {summary['両方']}件　/　"
        f"A のみ: {summary['A のみ']}件　/　B のみ: {summary['B のみ']}件"
    )
    st.dataframe(result['comparison'], width='stretch', hide_index=True)
    st.download_button(
        label="比較結果Excelをダウンロード",
        data=st.session_state['same_workbook_output'],
        file_name=f"label_comparison_UNIT内結線図_vs_{result['b_sheet_name']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type='primary',
        key='same_workbook_download',
    )


def _show_region_selection(file_a, file_b):
    """A・B両方に『領域別ラベル一覧』シートがある場合のみ、『指定領域で比較する』
    チェックボックスと共通領域名のマルチセレクトを表示する。

    戻り値: (region_mode: bool, selected_regions: list, a_region_rows, b_region_rows)
    region_mode が False の場合、残り3つは常に (False, [], None, None)。
    """
    a_bytes = file_a.getvalue()
    b_bytes = file_b.getvalue()
    if not (has_region_sheet(a_bytes) and has_region_sheet(b_bytes)):
        return False, [], None, None

    region_mode = st.checkbox(
        "指定領域で比較する", key='region_mode_enabled',
        help="A・B両方の『領域別ラベル一覧』シートを使い、選択した領域名だけを"
             "対象にラベルを比較します（Total シート全体の比較は行いません）。",
    )
    if not region_mode:
        return False, [], None, None

    try:
        a_region_rows = load_region_rows(a_bytes)
        b_region_rows = load_region_rows(b_bytes)
    except ValueError as error:
        st.error(f"領域データの読み込みエラー: {error}")
        return True, [], None, None

    a_region_names = {r[0] for r in a_region_rows}
    b_region_names = {r[0] for r in b_region_rows}
    common_regions = sorted(a_region_names & b_region_names)
    if not common_regions:
        st.warning("A・Bに共通する領域名がありません。")
        return True, [], a_region_rows, b_region_rows

    selected_regions = st.multiselect(
        "比較する領域名（A・Bに共通するものだけを表示・複数選択可）",
        common_regions, key='selected_regions',
    )
    return True, selected_regions, a_region_rows, b_region_rows


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

    region_mode, selected_regions, a_region_rows, b_region_rows = (False, [], None, None)
    if file_a is not None and file_b is not None:
        region_mode, selected_regions, a_region_rows, b_region_rows = _show_region_selection(
            file_a, file_b)

    run_disabled = file_a is None or file_b is None or (region_mode and not selected_regions)
    if st.button("展開図-結線図比較", type='primary', disabled=run_disabled,
                 key='two_workbook_run'):
        try:
            if region_mode:
                title_map = load_summary_titles(file_b.getvalue())
                selected_gzuban = select_drawing_numbers(title_map, filter_mode)
                a_by_region = aggregate_region_rows(a_region_rows, None)
                b_by_region = aggregate_region_rows(b_region_rows, selected_gzuban)
                diff_df, metrics_by_region = compare_labels_by_region(
                    a_by_region, b_by_region, selected_regions)
                summary = build_region_summary_rows(
                    metrics_by_region, file_a.name, file_b.name, b_filter_mode=filter_mode)
                st.session_state['two_workbook_output'] = create_region_compare_excel_output(
                    diff_df, summary)
                st.session_state['two_workbook_region_metrics'] = metrics_by_region
            else:
                labels_a = load_total_labels(file_a.getvalue())
                b_bytes = file_b.getvalue()
                title_map = load_summary_titles(b_bytes)
                total_rows_b = load_total_rows(b_bytes)
                selected_gzuban = select_drawing_numbers(title_map, filter_mode)
                labels_b = aggregate_filtered_rows(total_rows_b, selected_gzuban)
                diff_df = compare_labels(labels_a, labels_b)
                summary = summarize(diff_df, file_a.name, file_b.name, b_filter_mode=filter_mode)
                st.session_state['two_workbook_output'] = create_compare_excel_output(
                    diff_df, summary)
            st.session_state['two_workbook_diff'] = diff_df
            st.session_state['two_workbook_summary'] = summary
            st.session_state['two_workbook_is_region_mode'] = region_mode
        except ValueError as error:
            st.error(f"読み込みエラー: {error}")
            return

    diff_df = st.session_state.get('two_workbook_diff')
    if diff_df is None:
        return
    summary = st.session_state['two_workbook_summary']
    is_region_mode = st.session_state.get('two_workbook_is_region_mode', False)
    st.divider()
    st.subheader("結果")
    if is_region_mode:
        metrics_by_region = st.session_state.get('two_workbook_region_metrics', {})
        totals = {}
        for metrics in metrics_by_region.values():
            for key, value in metrics.items():
                totals[key] = totals.get(key, 0) + value
        st.info(
            f"対象領域: {len(metrics_by_region)}件　/　"
            f"A のみ: {totals.get('A のみ', 0)}件　/　B のみ: {totals.get('B のみ', 0)}件　/　"
            f"両方: {totals.get('両方', 0)}件"
        )
        st.dataframe(_region_summary_for_display(summary), width='stretch', hide_index=True)
    else:
        st.info(
            f"A のみ: {summary['A のみ']}件　/　B のみ: {summary['B のみ']}件　/　"
            f"両方: {summary['両方']}件"
        )
    st.dataframe(
        _for_display(diff_df).style.apply(_row_style_factory(diff_df), axis=1),
        width='stretch', hide_index=True,
    )
    st.download_button(
        label="Excelをダウンロード", data=st.session_state['two_workbook_output'],
        file_name="label_compare.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type='primary', key='two_workbook_download',
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
