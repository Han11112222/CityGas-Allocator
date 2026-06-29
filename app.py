import os
import re
import requests
import openpyxl
import pandas as pd
import streamlit as st
from io import BytesIO
from typing import Optional

# ──────────────────────────────────────────
# 설정 (GitHub 연동)
# ──────────────────────────────────────────
GITHUB_USER = "Han11112222"         # 👈 이 부분을 형님의 '실제 깃허브 아이디'로 변경!
GITHUB_REPO = "gas-settlement"       
GITHUB_BRANCH = "main"

# 정산 테이블에 표시할 표준 용도 순서 및 소계 여부 정의
TARGET_LABELS = [
    {"label": "주택용(소계)",   "is_subtotal": True},
    {"label": "일반용",         "is_subtotal": True},
    {"label": "냉난방공조용",   "is_subtotal": False},
    {"label": "업무난방용",     "is_subtotal": False},
    {"label": "산업용",         "is_subtotal": False},
    {"label": "수송용",         "is_subtotal": False},
    {"label": "열병합용",       "is_subtotal": False},
    {"label": "연료전지용",     "is_subtotal": False},
    {"label": "열전용설비용",   "is_subtotal": False},
    {"label": "주한미군",       "is_subtotal": False},
    {"label": "합계",           "is_subtotal": True},
]

# L열(판매량계) = 인덱스 11, M열(구성비) = 인덱스 12 (0-based)
COL_SALES_TOTAL = 11   # L열
COL_RATIO       = 12   # M열


# ──────────────────────────────────────────
# GitHub 파일 목록 & 다운로드
# ──────────────────────────────────────────
@st.cache_data(ttl=300)
def get_github_file_list() -> list[str]:
    """레포 루트의 파일 목록을 가져옴"""
    url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/git/trees/{GITHUB_BRANCH}?recursive=1"
    headers = {}
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        headers["Authorization"] = f"token {token}"
    r = requests.get(url, headers=headers, timeout=10)
    if r.status_code != 200:
        st.error(f"GitHub API 오류: {r.status_code} – {r.text[:200]}")
        return []
    tree = r.json().get("tree", [])
    return [item["path"] for item in tree if item["path"].endswith(".xlsx")]


@st.cache_data(ttl=300)
def download_xlsx(filename: str) -> Optional[bytes]:
    """raw 파일 다운로드"""
    url = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{requests.utils.quote(filename)}"
    headers = {}
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        headers["Authorization"] = f"token {token}"
    r = requests.get(url, headers=headers, timeout=15)
    if r.status_code == 200:
        return r.content
    return None


def parse_yearmonth(filename: str) -> Optional[tuple[int, int]]:
    """파일명에서 (연도, 월) 추출"""
    m = re.search(r"(\d{4})년(\d{1,2})월", filename)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


# ──────────────────────────────────────────
# 엑셀 동적 파싱 (하드코딩 행 번호 제거)
# ──────────────────────────────────────────
def read_gasco_sheet(xlsx_bytes: bytes) -> dict:
    """
    '가스공사용(MJ)' 시트에서 텍스트 매칭을 통해 용도별 판매량계(MJ)·구성비(%)를 동적 추출
    """
    wb = openpyxl.load_workbook(BytesIO(xlsx_bytes), read_only=True, data_only=True)
    if "가스공사용(MJ)" not in wb.sheetnames:
        return {}
    ws = wb["가스공사용(MJ)"]
    
    result = {}
    is_juteak_section = False  # 주택용 하부의 '소계'를 정확히 잡기 위한 플래그
    
    for row in ws.iter_rows(values_only=True):
        if len(row) <= max(COL_SALES_TOTAL, COL_RATIO):
            continue
            
        # B열(인덱스1)과 C열(인덱스2) 텍스트의 공백을 제거하여 비교 안정성 확보
        val_b = str(row[1]).replace(" ", "").strip() if row[1] is not None else ""
        val_c = str(row[2]).replace(" ", "").strip() if row[2] is not None else ""
        
        if "주택용" in val_b:
            is_juteak_section = True
            
        label = None
        is_subtotal = False
        
        # 텍스트 매칭 조건 분기
        if is_juteak_section and val_c == "소계":
            label = "주택용(소계)"
            is_subtotal = True
            is_juteak_section = False  # 주택용 구간 종료
        elif val_b == "일반용" and (val_c == "" or val_c == "소계"):
            label = "일반용"
            is_subtotal = True
        elif "냉난방공조용" in val_b or "냉난방용" in val_b:
            label = "냉난방공조용"
        elif "업무난방용" in val_b:
            label = "업무난방용"
        elif "산업용" in val_b:
            label = "산업용"
        elif "수송용" in val_b:
            label = "수송용"
        elif "열병합용" in val_b:
            label = "열병합용"
        elif "연료전지용" in val_b:
            label = "연료전지용"
        elif "열전용설비용" in val_b:
            label = "열전용설비용"
        elif "주한미군" in val_b:
            label = "주한미군"
        elif val_b in ["합계", "계"]:
            label = "합계"
            is_subtotal = True
            
        if label:
            sales = row[COL_SALES_TOTAL]
            ratio = row[COL_RATIO]
            if isinstance(sales, (int, float)) and sales:
                result[label] = {
                    "sales_mj": float(sales),
                    "ratio": float(ratio) if isinstance(ratio, (int, float)) else None,
                    "is_subtotal": is_subtotal,
                }
    return result


# ──────────────────────────────────────────
# 정산 계산
# ──────────────────────────────────────────
def calc_settlement(
    supply_gj: float,
    prev_data: dict,
    curr_data: dict,
    method: str = "당월단독",
) -> pd.DataFrame:
    """용도별 배분량(GJ) 계산"""
    rows = []

    total_prev = prev_data.get("합계", {}).get("sales_mj", None)
    total_curr = curr_data.get("합계", {}).get("sales_mj", None)

    for target in TARGET_LABELS:
        label = target["label"]
        prev = prev_data.get(label, {})
        curr = curr_data.get(label, {})

        prev_mj = prev.get("sales_mj")
        curr_mj = curr.get("sales_mj")

        # 구성비 계산 로직
        if method == "전월평균" and prev_mj and curr_mj and total_prev and total_curr:
            r_prev = prev_mj / total_prev * 100
            r_curr = curr_mj / total_curr * 100
            ratio = (r_prev + r_curr) / 2
        elif curr_mj and total_curr:
            ratio = curr_mj / total_curr * 100
        elif prev_mj and total_prev:
            ratio = prev_mj / total_prev * 100
        else:
            ratio = None

        alloc_gj = supply_gj * ratio / 100 if ratio is not None else None

        rows.append({
            "용도":           label,
            "구성비(%)":      round(ratio, 4) if ratio else None,
            "배분공급량(GJ)": round(alloc_gj, 3) if alloc_gj else None,
            "비고":           "소계" if target["is_subtotal"] else "",
        })

    return pd.DataFrame(rows)


# ──────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="도시가스 도매요금 정산 계산기",
        page_icon="🔥",
        layout="wide",
    )

    st.markdown(
        """
        <style>
        .main-title {
            font-size: 1.6rem; font-weight: 700;
            color: #1a3a5c; border-bottom: 2px solid #e65c00;
            padding-bottom: 0.3rem; margin-bottom: 0.2rem;
        }
        .sub-title { font-size: 0.9rem; color: #666; margin-bottom: 1.2rem; }
        .highlight { background: #f0f4fa; border-left: 3px solid #1a3a5c;
                     padding: 0.5rem 0.8rem; border-radius: 4px; font-size: 0.85rem; }
        </style>
        <div class="main-title">🔥 도시가스 도매요금 정산 계산기</div>
        <div class="sub-title">판매량정산서 기반 용도별 구성비 & 공급량 자동 산출</div>
        """,
        unsafe_allow_html=True,
    )

    with st.spinner("GitHub에서 판매량정산서 목록 로딩 중…"):
        xlsx_files = get_github_file_list()

    if not xlsx_files:
        st.warning("GitHub 레포에서 xlsx 파일을 찾지 못했습니다. GITHUB_USER / GITHUB_REPO 설정을 확인하거나 파일 확장자가 .xlsx인지 확인하세요.")
        return

    file_map: dict[tuple[int, int], str] = {}
    for f in xlsx_files:
        ym = parse_yearmonth(f)
        if ym:
            file_map[ym] = f

    if not file_map:
        st.error("파일명에 연도·월 패턴(예: 2025년08월)을 가진 xlsx 파일이 없습니다.")
        return

    sorted_keys = sorted(file_map.keys())
    month_labels = [f"{y}년 {m:02d}월" for y, m in sorted_keys]

    st.markdown("---")
    col1, col2, col3 = st.columns([2, 2, 2])

    with col1:
        st.markdown("**① 정산 대상 월 선택**")
        selected_label = st.selectbox("정산월", month_labels, index=len(month_labels) - 1)
        selected_ym = sorted_keys[month_labels.index(selected_label)]

    with col2:
        st.markdown("**② 수급량 입력 (GJ)**")
        supply_gj = st.number_input(
            "월 수급량 (GJ)",
            min_value=0.0,
            value=1774554.307,
            step=1000.0,
            format="%.3f",
            help="가스공사정산서 또는 공급량 보고서의 총 수급량(GJ)을 입력하세요.",
        )

    with col3:
        st.markdown("**③ 구성비 산출 방식**")
        method = st.radio(
            "방식",
            ["당월단독", "전월평균"],
            help="당월단독: 해당 월 판매량만 사용\n전월평균: (m-1)월 + m월 구성비 평균값 사용",
        )

    curr_idx = sorted_keys.index(selected_ym)
    prev_ym = sorted_keys[curr_idx - 1] if curr_idx > 0 else None

    st.markdown(
        f'<div class="highlight">📁 당월 파일: <b>{file_map[selected_ym]}</b>'
        + (f' &nbsp;|&nbsp; 전월 파일: <b>{file_map[prev_ym]}</b>' if prev_ym else "")
        + "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("")

    if st.button("🔢 구성비 & 공급량 산출", type="primary", use_container_width=True):
        with st.spinner("엑셀 파일 다운로드 및 정산 배분 계산 중…"):
            curr_bytes = download_xlsx(file_map[selected_ym])
            if not curr_bytes:
                st.error(f"당월 파일 다운로드 실패: {file_map[selected_ym]}")
                return
            curr_data = read_gasco_sheet(curr_bytes)

            prev_data = {}
            if prev_ym and method == "전월평균":
                prev_bytes = download_xlsx(file_map[prev_ym])
                if prev_bytes:
                    prev_data = read_gasco_sheet(prev_bytes)

            df = calc_settlement(supply_gj, prev_data, curr_data, method)

        st.markdown("### 📊 용도별 공급량 배분 결과")

        def style_row(row):
            if row["용도"] in ("합계", "주택용(소계)", "일반용"):
                return ["background-color: #f0f4fa; font-weight: bold"] * len(row)
            return [""] * len(row)

        main_df = df[df["용도"] != "합계"].copy()
        total_row = df[df["용도"] == "합계"].copy()

        styled = main_df.style.apply(style_row, axis=1).format({
            "구성비(%)":      lambda x: f"{x:.4f}%" if pd.notna(x) else "-",
            "배분공급량(GJ)": lambda x: f"{x:,.3f}" if pd.notna(x) else "-",
        })
        st.dataframe(styled, use_container_width=True, height=420)

        if not total_row.empty:
            t = total_row.iloc[0]
            c1, c2, c3 = st.columns(3)
            c1.metric("수급량 (입력값)", f"{supply_gj:,.3f} GJ")
            c2.metric("배분 총합계 (GJ)", f"{t['배분공급량(GJ)']:,.3f}" if pd.notna(t['배분공급량(GJ)']) else "-")
            c3.metric("구성비 총합", f"{t['구성비(%)']:.4f}%" if pd.notna(t['구성비(%)']) else "-")

        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label="⬇️ 정산 결과 CSV 다운로드",
            data=csv.encode("utf-8-sig"),
            file_name=f"도매정산결과_{selected_label.replace(' ', '')}.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
