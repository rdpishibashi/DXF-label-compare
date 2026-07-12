import streamlit as st

from utils.excel_input import load_total_labels
from utils.compare_labels import (
    compare_labels, summarize, KUBUN_BOTH, KUBUN_A_ONLY, KUBUN_B_ONLY,
)
from utils.excel_output import create_compare_excel_output

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


def main():
    st.title("DXF Label Compare")
    st.caption(
        "DXF-extract-labels が出力した2つのExcel（Totalシート）のラベルを比較し、"
        "差分（Aのみ・Bのみ・両方）をExcelで出力します。"
    )

    col_a, col_b = st.columns(2)
    with col_a:
        file_a = st.file_uploader(
            "A: 比較元の Excel（extract-labels 出力）", type=['xlsx'], key='uploader_a',
        )
    with col_b:
        file_b = st.file_uploader(
            "B: 比較先の Excel（extract-labels 出力）", type=['xlsx'], key='uploader_b',
        )

    has_input = file_a is not None and file_b is not None
    run = st.button("比較", type="primary", disabled=not has_input)

    if run:
        try:
            labels_a = load_total_labels(file_a.getvalue())
            labels_b = load_total_labels(file_b.getvalue())
        except ValueError as e:
            st.error(f"読み込みエラー: {e}")
            return

        diff_df = compare_labels(labels_a, labels_b)
        summary = summarize(diff_df, file_a.name, file_b.name)
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
        st.info(
            f"A のみ: {summary['A のみ']}件　/　"
            f"B のみ: {summary['B のみ']}件　/　"
            f"両方: {summary['両方']}件　/　"
            f"ユニーク合計: {summary['ユニーク合計']}件"
        )

        styled = diff_df.style.apply(_row_style, axis=1)
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
