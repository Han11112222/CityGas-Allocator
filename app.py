import os
import re
import requests
import pandas as pd
import streamlit as st
from io import BytesIO
from typing import Optional

# ──────────────────────────────────────────
# 설정 (GitHub 연동)
# ──────────────────────────────────────────
GITHUB_USER   = "Han11112222"
GITHUB_REPO   = "CityGas-Allocator"
GITHUB_BRANCH = "main"

TARGET_LABELS = [
    {"label": "주택용(소계)",  "is_subtotal": True},
    {"label": "일반용",        "is_subtotal": True},
    {"label": "냉난방공조용",  "is_subtotal": False},
    {"label": "업무난방용",    "is_subtotal": False},
    {"label": "산업용",        "is_subtotal": False},
    {"label": "수송용",        "is_subtotal": False},
    {"label": "열병합용",      "is_subtotal": False},
    {"label": "연료전지용",    "is_subtotal": False},
    {"label": "열전용설비용",  "is_subtotal": False},
    {"label": "주한미군",      "is_subtotal": False},
    {"label": "합계",          "is_subtotal": True},
]

COL_SALES_TOTAL = 11
COL_RATIO       = 12


# ──────────────────────────────────────────
# GitHub 파일 목록 & 다운로드
# ──────────────────────────────────────────
def get_github_file_list() -> list:
    url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/git/trees/{GITHUB_BRANCH}?recursive=1"
    headers = {}
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        headers["Authorization"] = f"token {token}"
    r = requests.get(url, headers=headers, timeout=10)

    st.session_state["api_status"] = r.status_code

    if r.status_code != 200:
        st.session_state["api_error_msg"] = r.text[:200]
        return []

    tree = r.json().get("tree", [])
    all_files = [item["path"] for item in tree]
    st.session_state["all_files"] = all_files
    return [p for p in all_files if p.endswith(".xlsx")]


@st.cache_data(ttl=300)
def download_xlsx(filename: str) -> Optional[bytes]:
    url = (
        f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}"
        f"/{GITHUB_BRANCH}/{requests.utils.quote(filename)}"
    )
    headers = {}
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        headers["Authorization"] = f"token {token}"
    r = requests.get(url, headers=headers, timeout=15)
    return r.content if r.status_code == 200 else None


def parse_yearmonth(filename: str) -> Optional[tuple]:
    m = re.search(r"(\d{4})년(\d{1,2})월", filename)
    return (int(m.group(1)), int(m.group(2))) if m else None


# ──────────────────────────────────────────
# 엑셀 파싱 (openpyxl 직접 호출 제거 → pandas 사용)
# ──────────────────────────────────────────
def read_gasco_sheet(xlsx_bytes: bytes) -> dict:
    try:
        xl = pd.ExcelFile(BytesIO(xlsx_bytes))
    except Exception as e:
        st.error(f"엑셀 파일 열기 실패: {e}")
        return {}

    if "가스공사용(MJ)" not in xl.sheet_names:
        return {}

    # header=None으로 전체를 raw 읽기
    df = xl.parse("가스공사용(MJ)", header=None)

    result = {}
    is_juteak_section = False

    for _, row in df.iterrows():
        # 컬럼 수 부족하면 스킵
        if len(row) <= max(COL_SALES_TOTAL, COL_RATIO):
            continue

        val_b = str(row.iloc[1]).replace(" ", "").strip() if row.iloc[1] is not None else ""
        val_c = str(row.iloc[2]).replace(" ", "").strip() if row.iloc[2] is not None else ""

        if "주택용" in val_b:
            is_juteak_section = True

        label = None
        is_subtotal = False

        if is_juteak_section and val_c == "소계":
            label, is_subtotal = "주택용(소계)", True
            is_juteak_section = False
        elif val_b in ["일반용", "일반용소계"] or (val_b == "일반용" and val_c in ["", "소계"]):
            label, is_subtotal = "일반용", True
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
            label, is_subtotal = "합계", True

        if label:
            sales = row.iloc[COL_SALES_TOTAL]
            ratio = row.iloc[COL_RATIO]
            try:
                sales_f = float(sales)
            except (TypeError, ValueError):
                sales_f = 0.0
            if sales_f:
                result[label] = {
                    "sales_mj":   sales_f,
                    "ratio":      float(ratio) if isinstance(ratio, (int, float)) else None,
                    "is_subtotal": is_subtotal,
                }
    return result


# ──────────────────────────────────────────
# 정산 계산
# ──────────────────────────────────────────
def calc_settlement(supply_gj, prev_data, curr_data, method="당월단독"):
    rows = []
    total_prev = prev_data.get("합계", {}).get("sales_mj")
    total_curr = curr_data.get("합계", {}).get("sales_mj")

    for target in TARGET_LABELS:
        label = target["label"]
        prev_mj = prev_data.get(label, {}).get("sales_mj")
        curr_mj = curr_data.get(label, {}).get("sales_mj")

        if method == "전월평균" and prev_mj and curr_mj and total_prev and total_curr:
            ratio = ((prev_mj / total_prev) + (curr_mj / total_curr)) / 2 * 100
        elif curr_mj and total_curr:
            ratio = curr_mj / total_curr * 100
        elif prev_mj and total_prev:
            ratio = prev_mj / total_prev * 100
        else:
            ratio = None

        alloc_gj = supply_gj * ratio / 100 if ratio is not None else None

        rows.append({
            "용도":           label,
            "구성비(%)":      round(ratio, 4) if ratio is not None else None,
            "배분공급량(GJ)": round(alloc_gj, 3) if alloc_gj is not None else None,
            "비고":           "소계" if target["is_subtotal"] else "",
        })
    return pd.DataFrame(rows)


# ──────────────────────────────────────────
# Streamlit UI
# ──────────────────────────────────────────
def main():
    st.set_page_config(page_title="도시가스 도매요금 정산", page_icon="🔥", layout="wide")

    st.markdown("""
        <style>
        .main-title { font-size:1.6rem; font-weight:700; color:#1a3a5c;
                      border-bottom:2px solid #e65c00; padding-bottom:.3rem; margin-bottom:.2rem; }
        .sub-title  { font-size:.9rem; color:#666; margin-bottom:1.2rem; }
        .highlight  { background:#f0f4fa; border-left:3px solid #1a3a5c;
                      padding:.5rem .8rem; border-radius:4px; font-size:.85rem; }
        </style>
        <div class="main-title">🔥 도시가스 도매요금 정산 계산기</div>
        <div class="sub-title">판매량정산서 기반 용도별 구성비 &amp; 공급량 자동 산출</div>
    """, unsafe_allow_html=True)

    # 파일 목록 로드
    with st.spinner("GitHub에서 판매량정산서 목록 로딩 중…"):
        xlsx_files = get_github_file_list()

    # 디버깅 패널
    with st.expander("🛠️ 시스템 상태 (문제 발생 시 확인)", expanded=not bool(xlsx_files)):
        status = st.session_state.get("api_status", "N/A")
        st.write(f"GitHub API 상태: `{status}` (200이면 정상)")
        if status != 200:
            st.error(f"오류: {st.session_state.get('api_error_msg', '')}")
        st.write("발견된 전체 파일:", st.session_state.get("all_files", []))

    if not xlsx_files:
        st.warning("xlsx 파일을 찾지 못했습니다. 위 디버깅 창을 확인해 주세요.")
        return

    # 파일 → (연도, 월) 매핑
    file_map = {}
    for f in xlsx_files:
        ym = parse_yearmonth(f)
        if ym:
            file_map[ym] = f

    if not file_map:
        st.error("연도·월 패턴(예: 2025년08월)을 가진 xlsx 파일이 없습니다.")
        return

    sorted_keys  = sorted(file_map.keys())
    month_labels = [f"{y}년 {m:02d}월" for y, m in sorted_keys]

    st.markdown("---")
    col1, col2, col3 = st.columns([2, 2, 2])

    with col1:
        st.markdown("**① 정산 대상 월 선택**")
        selected_label = st.selectbox("정산월", month_labels, index=len(month_labels) - 1)
        selected_ym    = sorted_keys[month_labels.index(selected_label)]

    with col2:
        st.markdown("**② 수급량 입력 (GJ)**")
        supply_gj = st.number_input(
            "월 수급량 (GJ)", min_value=0.0, value=1_774_554.307,
            step=1000.0, format="%.3f",
            help="가스공사 공급량 일일보고서의 월 수급량(GJ)"
        )

    with col3:
        st.markdown("**③ 구성비 산출 방식**")
        method = st.radio("방식", ["당월단독", "전월평균"],
                          help="당월단독: 해당 월만 사용 / 전월평균: (m-1)월+m월 평균")

    curr_idx = sorted_keys.index(selected_ym)
    prev_ym  = sorted_keys[curr_idx - 1] if curr_idx > 0 else None

    st.markdown(
        f'<div class="highlight">📁 당월: <b>{file_map[selected_ym]}</b>'
        + (f' &nbsp;|&nbsp; 전월: <b>{file_map[prev_ym]}</b>' if prev_ym else "")
        + "</div>", unsafe_allow_html=True
    )
    st.markdown("")

    if st.button("🔢 구성비 & 공급량 산출", type="primary", use_container_width=True):
        with st.spinner("엑셀 파일 다운로드 & 계산 중…"):
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
                return ["background-color:#f0f4fa; font-weight:bold"] * len(row)
            return [""] * len(row)

        main_df  = df[df["용도"] != "합계"].copy()
        total_row = df[df["용도"] == "합계"].copy()

        st.dataframe(
            main_df.style.apply(style_row, axis=1).format({
                "구성비(%)":      lambda x: f"{x:.4f}%" if pd.notna(x) else "-",
                "배분공급량(GJ)": lambda x: f"{x:,.3f}" if pd.notna(x) else "-",
            }),
            use_container_width=True, height=420
        )

        if not total_row.empty:
            t = total_row.iloc[0]
            c1, c2, c3 = st.columns(3)
            c1.metric("수급량 (입력)", f"{supply_gj:,.3f} GJ")
            c2.metric("배분 합계 (GJ)",
                      f"{t['배분공급량(GJ)']:,.3f}" if pd.notna(t["배분공급량(GJ)"]) else "-")
            c3.metric("구성비 합계",
                      f"{t['구성비(%)']:.4f}%" if pd.notna(t["구성비(%)"]) else "-")

        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label="⬇️ 정산 결과 CSV 다운로드",
            data=csv.encode("utf-8-sig"),
            file_name=f"도매정산_{selected_label.replace(' ', '')}.csv",
            mime="text/csv",
        )

    with st.expander("ℹ️ 로직 설명"):
        st.markdown("""
| 단계 | 내용 |
|------|------|
| ① | GitHub에서 판매량정산서 xlsx 자동 다운로드 |
| ② | `가스공사용(MJ)` 시트 → 용도별 판매량계(MJ) 추출 |
| ③ | 판매량계 / 합계 × 100 = 용도별 구성비(%) |
| ④ | 구성비 × 수급량(GJ) = 용도별 배분 공급량(GJ) |
        """)


if __name__ == "__main__":
    main()
