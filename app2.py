import sys
import importlib.util

# ---------------------------------------------------------
# 0. openpyxl 자동 설치 및 체크 (내장 모듈 활용)
# ---------------------------------------------------------
if importlib.util.find_spec("openpyxl") is None:
    import subprocess

    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

import openpyxl
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import io

# ---------------------------------------------------------
# 1. 페이지 설정
# ---------------------------------------------------------
st.set_page_config(
    page_title="진급연한 현황 대시보드",
    page_icon="📋",
    layout="wide"
)

# ---------------------------------------------------------
# 2. 세션 상태(Session State) 초기화
# ---------------------------------------------------------
if "today_date" not in st.session_state:
    st.session_state.today_date = date.today()

if "filter_status" not in st.session_state:
    st.session_state.filter_status = "all"

if "df_data" not in st.session_state:
    st.session_state.df_data = None


# 오늘 날짜로 리셋 콜백
def reset_date_callback():
    st.session_state.today_date = date.today()
    st.session_state.date_picker = date.today()


# ---------------------------------------------------------
# 3. 커스텀 CSS (파스텔톤 UI & 스타일링)
# ---------------------------------------------------------
st.markdown("""
<style>
    .promotion-dashboard-scope {
        font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "Pretendard", "맑은 고딕", sans-serif;
    }

    .main-title {
        font-size: 24px;
        font-weight: 800;
        color: #111111;
        margin-bottom: 5px;
    }

    .custom-table-container {
        width: 100%;
        overflow-x: auto;
        margin-top: 15px;
    }

    table.promotion-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        font-size: 13px;
        font-family: -apple-system, BlinkMacSystemFont, "Apple SD Gothic Neo", "맑은 고딕", sans-serif !important;
    }

    table.promotion-table th {
        background-color: #F8FAFC;
        color: #334155;
        font-weight: 700;
        padding: 10px 8px;
        text-align: center !important;
        border-bottom: 2px solid #E2E8F0;
        border-right: 1px solid #E2E8F0;
    }

    table.promotion-table td {
        padding: 9px 8px;
        text-align: center !important;
        vertical-align: middle;
        border-bottom: 1px solid #F1F5F9;
        border-right: 1px solid #F1F5F9;
        background-color: #FFFFFF;
        white-space: nowrap;
        font-weight: 500;
        color: #1E293B;
    }

    table.promotion-table tr:hover td {
        background-color: #F8FAFC;
    }

    /* 파스텔톤 강조 셀 서식 */
    .cell-expired {
        background-color: #FEF2F2 !important;
        color: #DC2626 !important;
        font-weight: 700 !important;
        border: 1.5px solid #FECACA !important;
        border-radius: 4px;
    }

    .cell-within-3m {
        background-color: #EFF6FF !important;
        color: #2563EB !important;
        font-weight: 700 !important;
        border: 1.5px solid #BFDBFE !important;
        border-radius: 4px;
    }

    .cell-within-6m {
        background-color: #ECFDF5 !important;
        color: #059669 !important;
        font-weight: 700 !important;
        border: 1.5px solid #A7F3D0 !important;
        border-radius: 4px;
    }

    .cell-next-year {
        background-color: #F5F3FF !important;
        color: #7C3AED !important;
        font-weight: 700 !important;
        border: 1.5px solid #DDD6FE !important;
        border-radius: 4px;
    }

    .cell-mismatch {
        background-color: #FFFBEB !important;
        color: #D97706 !important;
        font-weight: 700 !important;
        border: 1.5px solid #FDE68A !important;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="promotion-dashboard-scope main-title">📋 진급연한 현황 대시보드</div>', unsafe_allow_html=True)
st.markdown("---")

# ---------------------------------------------------------
# 4. 예시 데이터 생성 및 기본 로드
# ---------------------------------------------------------
REQUIRED_COLS = ["사번", "이름", "학력", "입사일", "부서", "직위", "사원", "주임", "대리", "과장", "차장", "부장", "비고"]


def generate_sample_excel():
    data = [
        {
            "사번": "2020001", "이름": "홍길동", "학력": "대졸", "입사일": "2020-03-01",
            "부서": "영업팀", "직위": "대리",
            "사원": "2020-03-01", "주임": "2022-03-01", "대리": "2024-03-01",
            "과장": "2026-03-01", "차장": "", "부장": "",
            "비고": "진급일 경과 예시"
        },
        {
            "사번": "2021002", "이름": "김철수", "학력": "대졸", "입사일": "2021-01-15",
            "부서": "개발팀", "직위": "주임",
            "사원": "2021-01-15", "주임": "2023-01-15", "대리": "2026-11-15",
            "과장": "", "차장": "", "부장": "",
            "비고": "3개월 이내 도래 예시"
        },
        {
            "사번": "2022005", "이름": "박민수", "학력": "대졸", "입사일": "2022-03-01",
            "부서": "인사팀", "직위": "사원",
            "사원": "2022-03-01", "주임": "2027-03-01", "대리": "",
            "과장": "", "차장": "", "부장": "",
            "비고": "내년(2027년) 도래 예시"
        },
        {
            "사번": "2018003", "이름": "이영희", "학력": "석사", "입사일": "2018-07-01",
            "부서": "기획팀", "직위": "대리",
            "사원": "2018-07-01", "주임": "2020-07-01", "대리": "2022-07-01",
            "과장": "2025-07-01", "차장": "", "부장": "",
            "비고": "직급연한 미부합 예시"
        }
    ]
    df_ex = pd.DataFrame(data)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_ex.to_excel(writer, index=False, sheet_name='진급연한')
    return out.getvalue()


# ---------------------------------------------------------
# 5. 사이드바 - 설정 및 리셋
# ---------------------------------------------------------
st.sidebar.header("📂 데이터 파일 관리")
sample_excel = generate_sample_excel()
st.sidebar.download_button(
    label="📥 표준 양식 엑셀 파일 다운로드",
    data=sample_excel,
    file_name="진급연한_양식.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.sidebar.markdown("---")
uploaded_file = st.sidebar.file_uploader("엑셀/CSV 파일 업로드", type=["xlsx", "xls", "csv"])

# 파일 업로드 세션 보존
if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        raw_df = pd.read_csv(uploaded_file)
    else:
        raw_df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.session_state.df_data = raw_df
elif st.session_state.df_data is None:
    st.session_state.df_data = pd.read_excel(io.BytesIO(sample_excel), engine='openpyxl')

# 초기화 버튼
if st.sidebar.button("🗑️ 업로드 데이터 초기화"):
    st.session_state.df_data = pd.read_excel(io.BytesIO(sample_excel), engine='openpyxl')
    st.session_state.filter_status = "all"
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("📅 기준일 (TODAY) 설정")

selected_date_input = st.sidebar.date_input(
    "기준일 선택",
    value=st.session_state.today_date,
    key="date_picker"
)
st.session_state.today_date = selected_date_input

st.sidebar.button("🔄 오늘 날짜로 리셋", on_click=reset_date_callback, use_container_width=True)

# ---------------------------------------------------------
# 6. 데이터 전처리
# ---------------------------------------------------------
df = st.session_state.df_data.copy()

for col in REQUIRED_COLS:
    if col not in df.columns:
        df[col] = ""
df = df[REQUIRED_COLS]

DATE_COLS = ["입사일", "사원", "주임", "대리", "과장", "차장", "부장"]
for col in DATE_COLS:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d').fillna("")

df = df.fillna("")

# ---------------------------------------------------------
# 7. 개별 행 다중 조건 세트 측정 (필터링용)
# ---------------------------------------------------------
RANK_ORDER = ["사원", "주임", "대리", "과장", "차장", "부장"]
today_dt = pd.to_datetime(st.session_state.today_date)
in_3m = today_dt + relativedelta(months=3)
in_6m = today_dt + relativedelta(months=6)
next_year_num = today_dt.year + 1


def analyze_row_conditions(row):
    tags = set()
    current_rank = str(row.get("직위", "")).strip()

    # 1. 직급 불일치 확인
    next_rank = None
    if current_rank in RANK_ORDER:
        curr_idx = RANK_ORDER.index(current_rank)
        if curr_idx + 1 < len(RANK_ORDER):
            next_rank = RANK_ORDER[curr_idx + 1]

    if next_rank and next_rank in DATE_COLS:
        next_date_str = str(row.get(next_rank, "")).strip()
        if next_date_str:
            p_next_date = pd.to_datetime(next_date_str, errors='coerce')
            if pd.notnull(p_next_date) and p_next_date < today_dt:
                tags.add("mismatch")

    # 2. 직급 연한 날짜 범위별 조건 확인
    for col in RANK_ORDER:
        cell_value = str(row.get(col, "")).strip()
        if cell_value:
            p_date = pd.to_datetime(cell_value, errors='coerce')
            if pd.notnull(p_date):
                if p_date < today_dt:
                    tags.add("expired")
                elif today_dt <= p_date <= in_3m:
                    tags.add("within_3m")
                elif in_3m < p_date <= in_6m:
                    tags.add("within_6m")
                elif p_date.year == next_year_num:
                    tags.add("next_year")

    return tags


df['__tags'] = df.apply(analyze_row_conditions, axis=1)

# 카운트 집계
counts = {
    "all": len(df),
    "expired": sum(df['__tags'].apply(lambda x: "expired" in x)),
    "within_3m": sum(df['__tags'].apply(lambda x: "within_3m" in x)),
    "within_6m": sum(df['__tags'].apply(lambda x: "within_6m" in x)),
    "next_year": sum(df['__tags'].apply(lambda x: "next_year" in x)),
    "mismatch": sum(df['__tags'].apply(lambda x: "mismatch" in x)),
}

# ---------------------------------------------------------
# 8. 상단 필터 버튼 (클릭 시 실시간 동기화)
# ---------------------------------------------------------
col1, col2, col3, col4, col5, col6 = st.columns(6)


def set_filter(filter_name):
    st.session_state.filter_status = filter_name


with col1:
    st.button(f"🌐 전체 보기\n({counts['all']}명)", on_click=set_filter, args=("all",), use_container_width=True)
with col2:
    st.button(f"🔴 경과됨\n({counts['expired']}명)", on_click=set_filter, args=("expired",), use_container_width=True)
with col3:
    st.button(f"🔵 3개월 이내\n({counts['within_3m']}명)", on_click=set_filter, args=("within_3m",), use_container_width=True)
with col4:
    st.button(f"🟢 3~6개월 이내\n({counts['within_6m']}명)", on_click=set_filter, args=("within_6m",),
              use_container_width=True)
with col5:
    st.button(f"🟣 내년({next_year_num}년) 도래\n({counts['next_year']}명)", on_click=set_filter, args=("next_year",),
              use_container_width=True)
with col6:
    st.button(f"🟡 연한 불일치\n({counts['mismatch']}명)", on_click=set_filter, args=("mismatch",), use_container_width=True)

# 데이터 필터링 적용
if st.session_state.filter_status == "all":
    filtered_df = df.copy()
else:
    filtered_df = df[df['__tags'].apply(lambda tags: st.session_state.filter_status in tags)].copy()

filter_names = {
    "all": "전체 보기", "expired": "경과됨", "within_3m": "3개월 이내",
    "within_6m": "3~6개월 이내", "next_year": f"내년({next_year_num}년) 도래", "mismatch": "연한 불일치"
}
st.caption(f"📌 현재 선택된 필터: **{filter_names.get(st.session_state.filter_status, '전체')}** (총 {len(filtered_df)}명)")


# ---------------------------------------------------------
# 9. HTML 테이블 생성 함수
# ---------------------------------------------------------
def build_html_table(data_frame):
    table_html = '<div class="custom-table-container"><table class="promotion-table"><thead><tr>'

    display_cols = [c for c in REQUIRED_COLS if c in data_frame.columns]
    for col in display_cols:
        table_html += f'<th>{col}</th>'
    table_html += '</tr></thead><tbody>'

    for idx, row in data_frame.iterrows():
        table_html += '<tr>'
        current_rank = str(row.get("직위", "")).strip()

        next_rank = None
        if current_rank in RANK_ORDER:
            curr_idx = RANK_ORDER.index(current_rank)
            if curr_idx + 1 < len(RANK_ORDER):
                next_rank = RANK_ORDER[curr_idx + 1]

        is_mismatch = False
        if next_rank and next_rank in DATE_COLS:
            next_date_str = str(row.get(next_rank, "")).strip()
            if next_date_str:
                p_next_date = pd.to_datetime(next_date_str, errors='coerce')
                if pd.notnull(p_next_date) and p_next_date < today_dt:
                    is_mismatch = True

        for col in display_cols:
            cell_value = str(row[col]).strip()
            cell_class = ""

            if is_mismatch and (col in ["직위", "이름", next_rank]):
                cell_class = "cell-mismatch"
            elif col in RANK_ORDER and cell_value:
                p_date = pd.to_datetime(cell_value, errors='coerce')
                if pd.notnull(p_date):
                    if p_date < today_dt:
                        cell_class = "cell-expired"
                    elif today_dt <= p_date <= in_3m:
                        cell_class = "cell-within-3m"
                    elif in_3m < p_date <= in_6m:
                        cell_class = "cell-within-6m"
                    elif p_date.year == next_year_num:
                        cell_class = "cell-next-year"

            class_attr = f' class="{cell_class}"' if cell_class else ''
            table_html += f'<td{class_attr}>{cell_value}</td>'

        table_html += '</tr>'

    table_html += '</tbody></table></div>'
    return table_html


# ---------------------------------------------------------
# 10. 화면 출력 및 CSV 다운로드
# ---------------------------------------------------------
st.markdown(build_html_table(filtered_df), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

csv = filtered_df[REQUIRED_COLS].to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 현재 화면 데이터 CSV 다운로드",
    data=csv,
    file_name=f"진급연한_현황_{today_dt.strftime('%Y%m%d')}.csv",
    mime="text/csv"
)
