"""학점도"""

from __future__ import annotations

import html

import pandas as pd
import streamlit as st

from clean_all_records import process_workbook
from similarity_review import (
    KIND_COLORS,
    KIND_LABELS,
    build_review_excel,
    colored_content_html,
    records_from_processed_items,
    review_all_files,
    review_table,
)

st.set_page_config(
    page_title="학교생활기록부 작성 내용 검토",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)


def render_legend() -> None:
    st.markdown(
        f"""
        <div style="line-height:1.9">
        <strong>검토 결과 색상:</strong>
        <span style="color:{KIND_COLORS['exact']};font-weight:700">■ {KIND_LABELS['exact']}</span>&nbsp;&nbsp;
        <span style="color:{KIND_COLORS['sequence']};font-weight:700">■ {KIND_LABELS['sequence']}</span>&nbsp;&nbsp;
        <span style="color:{KIND_COLORS['similar']};font-weight:700">■ {KIND_LABELS['similar']}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_colored_preview_html(results: list[object], cross: bool) -> str:
    """검토 결과 원문과 판정 기준을 한 표 안에서 색상으로 직관적으로 표시한다."""
    number_header = "교차 동일 문장 사용 학생 번호" if cross else "동일 문장 사용 학생 번호"
    rows: list[str] = []
    for result in results:
        if not result.has_match():
            continue
        labels = "<br>".join(
            f'<span style="color:{KIND_COLORS[kind]};font-weight:700">● {html.escape(KIND_LABELS[kind])}</span>'
            for kind in result.kinds()
        )
        student_numbers = result.cross_student_numbers() if cross else result.local_student_numbers()
        rows.append(
            "<tr>"
            f"<td>{html.escape(result.record.grade_class)}</td>"
            f"<td style=\"text-align:center\">{html.escape(result.record.number)}번</td>"
            f"<td>{html.escape(result.record.category)}</td>"
            f"<td style=\"min-width:500px;line-height:1.75\">{colored_content_html(result)}</td>"
            f"<td>{labels}</td>"
            f"<td>{html.escape(student_numbers)}</td>"
            "</tr>"
        )
    if not rows:
        return ""
    return (
        '<div style="overflow-x:auto">'
        '<table style="width:100%;border-collapse:collapse;font-size:0.92rem">'
        '<thead><tr style="background:#1F4E78;color:#FFFFFF">'
        '<th style="padding:9px;border:1px solid #D9E2F3">학년 반</th>'
        '<th style="padding:9px;border:1px solid #D9E2F3">번호</th>'
        '<th style="padding:9px;border:1px solid #D9E2F3">검토 대상 영역</th>'
        '<th style="padding:9px;border:1px solid #D9E2F3">검토 내용 색상 미리보기</th>'
        '<th style="padding:9px;border:1px solid #D9E2F3">검출 기준</th>'
        f'<th style="padding:9px;border:1px solid #D9E2F3">{html.escape(number_header)}</th>'
        '</tr></thead><tbody>'
        + ''.join(rows)
        + '</tbody></table></div>'
    )


def render_review_tab(
    results: list[object],
    tab_key: str,
    title: str,
    cross: bool = False,
    base_file: str | None = None,
) -> None:
    """파일별 또는 교차 검토 결과를 표와 문장별 색상 미리보기로 표시한다."""
    st.subheader(title)
    if cross and base_file:
        st.info("교차 검토 결과는 먼저 업로드한 파일을 기준으로 표시됩니다.")
    elif not cross:
        st.caption("같은 파일 안에서 동일한 과목 또는 활동 영역의 서로 다른 학생 전체 내용을 비교합니다.")

    render_legend()
    matches_only = st.checkbox(
        "동일성 검출 행만 표시",
        value=True,
        key=f"matches_only_{tab_key}",
    )
    filtered_results = [result for result in results if result.has_match()] if matches_only else results
    match_count = sum(result.has_match() for result in results)
    st.caption(f"검토 대상 {len(results):,}행 중 동일성 검출 행은 **{match_count:,}행**입니다.")

    table = review_table(results, matches_only=matches_only, cross=cross)
    if match_count == 0:
        st.info(
            f"학생 {len(results):,}행의 비교 분석은 완료되었습니다. 현재 완전 동일, 연속 3문장 동일, "
            "전체 내용 유사도 95% 이상 기준을 충족한 행은 없습니다."
        )
        if matches_only:
            st.caption("`동일성 검출 행만 표시`를 해제하면 분석된 전체 학생 행을 확인할 수 있습니다.")
        else:
            st.dataframe(table, use_container_width=True, hide_index=True)
        return

    st.markdown("#### 검토 결과 색상 미리보기")
    st.markdown(build_colored_preview_html(filtered_results, cross=cross), unsafe_allow_html=True)

    with st.expander("표 형태의 상세 데이터 보기", expanded=False):
        st.dataframe(table, use_container_width=True, hide_index=True)


st.title("학교생활기록부 내용 동일성 검토")
st.markdown(
    """
업로드한 학교생활기록부 엑셀 파일을 자동 정제한 뒤 학생별 전체 내용을 비교합니다.
완전 동일, 붙어 있는 연속 3문장 이상 동일, 전체 내용 유사도 95% 이상을 색상으로 구분하여 검토합니다.
"""
)

with st.expander("동일성 검토 기준", expanded=False):
    st.markdown(
        """
        **전체 내용 완전 동일**은 두 학생의 모든 문장이 같은 경우입니다. **연속 3문장 이상 동일**은 두 학생의 내용 중
        순서가 유지된 붙어 있는 세 문장 이상이 같은 경우입니다. **전체 내용 유사도 95% 이상**은 조사, 부사, 활동명 등의
        일부 표현만 달라 전체 문자 배열 유사도가 95% 이상인 경우입니다. 서로 다른 과목 또는 활동 영역의 내용은 비교하지 않습니다.
        """
    )

uploaded_files = st.file_uploader(
    "학교생활기록부 엑셀 파일 업로드 (.xlsx, 여러 파일 선택 가능)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="파일 선택 순서가 교차 검토의 기준 파일 순서입니다.",
)

if not uploaded_files:
    st.info("동일성 검토를 할 엑셀 파일을 업로드해 주세요.")
    st.stop()

processed_items: list[dict[str, object]] = []
upload_order = [uploaded_file.name for uploaded_file in uploaded_files]

with st.spinner("학생별 내용을 정제하고 동일성을 검토하고 있습니다..."):
    for uploaded_file in uploaded_files:
        try:
            processed, _ = process_workbook(uploaded_file, source_name=uploaded_file.name)
            processed_items.extend(processed)
        except Exception:
            # 화면을 검토 결과 중심으로 유지하며, 정제 가능한 파일만 분석에 사용한다.
            continue

if not processed_items:
    st.error("검토 가능한 학생 데이터가 없습니다. 원본 엑셀의 제목과 헤더 양식을 확인해 주세요.")
    st.stop()

records_by_file = records_from_processed_items(processed_items)
file_results, cross_results, base_file = review_all_files(records_by_file, upload_order)

st.download_button(
    label="색상 포함 검토 결과 다운로드 (.xlsx)",
    data=build_review_excel(file_results, cross_results, matches_only=True),
    file_name="school_record_similarity_review.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    help="동일성이 검출된 행만 포함하며, 원본 파일명과 시트 열은 표시하지 않습니다.",
)

# 업로드 파일별 검토 탭과 기준 파일 우선 교차 검토 탭만 표시한다.
tab_file_order = [source_file for source_file in upload_order if source_file in file_results]
tab_labels = [f"파일 검토 · {source_file}" for source_file in tab_file_order] + ["교차 검토"]
tabs = st.tabs(tab_labels)

for tab_index, source_file in enumerate(tab_file_order):
    with tabs[tab_index]:
        render_review_tab(
            file_results[source_file],
            tab_key=f"file_{tab_index}",
            title=f"파일별 동일성 검토 — {source_file}",
            cross=False,
        )

with tabs[-1]:
    if len(tab_file_order) < 2:
        st.warning("교차 검토는 정제 가능한 엑셀 파일을 두 개 이상 업로드해야 실행됩니다.")
    else:
        render_review_tab(
            cross_results,
            tab_key="cross",
            title="교차 동일성 검토",
            cross=True,
            base_file=base_file,
        )
