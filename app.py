import os
import re
import requests
import pandas as pd
import streamlit as st
from io import BytesIO
from typing import Optional

# ──────────────────────────────────────────
# 설정
# ──────────────────────────────────────────
GITHUB_USER   = "Han11112222"
GITHUB_REPO   = "CityGas-Allocator"
GITHUB_BRANCH = "main"

# '3월검증' 형식 시트(상품/서비스/N월보정공급량) → 가스공사 9개 용도 매핑
# (상품, 서비스) 튜플로 정확히 매칭. 수송용(BIO)는 의도적으로 제외.
SERVICE_MAP = {
    "주택용": [
        ("개별난방용", "개별난방"),
        ("난방용", "개별난방"),
        ("냉난방용(주택)", "중앙난방"),
        ("자가열전용", "자가열전용"),
        ("중앙난방용", "중앙난방"),
        ("취사난방용", "개별난방"),
        ("취사난방용", "개별난방(취사)"),
        ("취사용", "취사용"),
    ],
    "업무난방용": [
        ("업무난방용", "업무난방"),
    ],
    "일반용": [
        ("일반용(1)", "일반용(1)(기타)"),
        ("일반용(1)", "일반용(1)(동절기)"),
        ("일반용(2)", "일반용(2)(기타)"),
        ("일반용(2)", "일반용(2)(동절기)"),
    ],
    "냉난방공조용": [
        ("냉난방용(업무)", "냉난방(기타)"),
        ("냉난방용(업무)", "냉난방(동절기)"),
    ],
    "산업용": [
        ("산업용", "산업용(기타)"),
        ("산업용", "산업용(동절기)"),
    ],
    "수송용": [
        ("수송용(CNG)", "수송용(CNG)"),
        ("수송용(외주)", "수송용(외주)"),
        # ("수송용(BIO)", "수송용(BIO)"),  ← 의도적 제외: 별도 바이오가스 정산
    ],
    "열병합용": [
        ("열병합용", "열병합용(기타)"),
        ("열병합용", "열병합용(동절기)"),
    ],
    "연료전지용": [
        ("연료전지", "연료전지(기타)"),
        ("연료전지", "연료전지(동절기)"),
    ],
    "열전용설비용": [
        ("열전용설비용", "열전용설비용"),
    ],
    "주한미군": [
        ("주한미군", "주한미군"),
    ],
}

TARGET_ORDER = [
    "주택용", "업무난방용", "일반용", "냉난방공조용", "산업용",
    "열병합용", "연료전지용", "열전용설비용", "수송용", "주한미군",
]
SUBTOTAL_LABELS = {"주택용", "일반용"}


# ──────────────────────────────────────────
# GitHub 연동
# ──────────────────────────────────────────
def get_github_file_list() -> list:
    url = (f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}"
           f"/git/trees/{GITHUB_BRANCH}?recursive=1")
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
    url = (f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}"
           f"/{GITHUB_BRANCH}/{requests.utils.quote(filename)}")
    headers = {}
    token = os.environ.get("GITHUB_TOKEN", "")
    if token:
        headers["Authorization"] = f"token {token}"
    r = requests.get(url, headers=headers, timeout=15)
    return r.content if r.status_code == 200 else None


def parse_target_month_from_filename(filename: str) -> Optional[tuple]:
    """파일명에서 기준연월 추정 (예: ...전체_3월_추정.xlsx → 해당 연도 3월)
       연도 정보가 없으면 None 반환, UI에서 직접 매칭"""
    m = re.search(r"(\d{1,2})월", filename)
    return int(m.group(1)) if m else None


# ──────────────────────────────────────────
# '3월검증' 형식 시트 파싱 (상품/서비스/N월보정공급량 구조)
# ──────────────────────────────────────────
def read_verification_sheet(xlsx_bytes: bytes, sheet_name: str = "3월검증") -> Optional[pd.DataFrame]:
    """
    원본 raw data 시트를 읽어 (상품, 서비스, 2월보정, 3월보정, 4월보정) DataFrame 반환
    헤더 행(9행)을 자동 탐지하여 '상품'/'서비스' 컬럼을 찾는다.
    """
    try:
        xl = pd.ExcelFile(BytesIO(xlsx_bytes))
    except Exception as e:
        st.error(f"엑셀 파일 열기 실패: {e}")
        return None

    if sheet_name not in xl.sheet_names:
        # 시트명이 다를 수 있으므로 '검증'이 포함된 시트를 탐색
        candidates = [s for s in xl.sheet_names if "검증" in s]
        if not candidates:
            return None
        sheet_name = candidates[0]

    raw = xl.parse(sheet_name, header=None)

    # '상품' 헤더가 있는 행 탐색
    header_row_idx = None
    for i in range(min(20, len(raw))):
        row_vals = [str(v).strip() for v in raw.iloc[i].tolist()]
        if "상품" in row_vals and "서비스" in row_vals:
            header_row_idx = i
            break
    if header_row_idx is None:
        return None

    header = raw.iloc[header_row_idx].tolist()
    data_rows = raw.iloc[header_row_idx + 1:].reset_index(drop=True)
    data_rows.columns = header

    # 보정공급량 컬럼 자동 탐지 (예: '합계 : 2월보정공급량')
    month_cols = {}
    for col in data_rows.columns:
        col_str = str(col)
        m = re.search(r"(\d{1,2})월보정공급량", col_str)
        if m and "합계" in col_str:
            month_cols[int(m.group(1))] = col

    # 상품/서비스 forward-fill (병합 셀 대응)
    data_rows["상품"] = data_rows["상품"].ffill()

    result_cols = ["상품", "서비스"] + list(month_cols.values())
    df = data_rows[result_cols].copy()
    df = df.dropna(subset=["서비스"])
    df = df[df["상품"] != "총합계"]

    # 컬럼명을 월 숫자로 정규화
    rename_map = {v: f"M{k}" for k, v in month_cols.items()}
    df = df.rename(columns=rename_map)

    for c in df.columns:
        if c.startswith("M"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    return df


def allocate_by_usage(df: pd.DataFrame, month_col: str) -> dict:
    """
    SERVICE_MAP에 따라 (상품,서비스) 쌍을 가스공사 9개 용도로 합산.
    df: read_verification_sheet 결과, month_col: 'M3' 등
    """
    if df is None or month_col not in df.columns:
        return {}

    # (상품, 서비스) → 값 딕셔너리
    lookup = {}
    for _, row in df.iterrows():
        key = (str(row["상품"]).strip(), str(row["서비스"]).strip())
        lookup[key] = lookup.get(key, 0) + row[month_col]

    result = {}
    for label, pairs in SERVICE_MAP.items():
        total = sum(lookup.get(pair, 0) for pair in pairs)
        result[label] = total

    return result


# ──────────────────────────────────────────
# 정산 계산
# ──────────────────────────────────────────
def calc_settlement(supply_gj: float, usage_totals: dict) -> pd.DataFrame:
    grand_total = sum(usage_totals.values())
    rows = []
    for label in TARGET_ORDER:
        val = usage_totals.get(label, 0)
        ratio = (val / grand_total * 100) if grand_total > 0 else None
        alloc_gj = supply_gj * ratio / 100 if ratio is not None else None
        rows.append({
            "용도":           label,
            "보정공급량(원본단위)": round(val, 0),
            "구성비(%)":      round(ratio, 4) if ratio is not None else None,
            "배분공급량(GJ)": round(alloc_gj, 3) if alloc_gj is not None else None,
        })
    total_row = {
        "용도": "합계",
        "보정공급량(원본단위)": round(grand_total, 0),
        "구성비(%)": 100.0 if grand_total > 0 else None,
        "배분공급량(GJ)": round(supply_gj, 3) if grand_total > 0 else None,
    }
    rows.append(total_row)
    return pd.DataFrame(rows)


# ──────────────────────────────────────────
# UI
# ──────────────────────────────────────────
def main():
    st.set_page_config(page_title="도시가스 도매요금 정산 (Raw Data 기반)", page_icon="🔥", layout="wide")

    st.markdown("""
        <style>
        .main-title { font-size:1.6rem; font-weight:700; color:#1a3a5c;
                      border-bottom:2px solid #e65c00; padding-bottom:.3rem; margin-bottom:.2rem; }
        .sub-title  { font-size:.9rem; color:#666; margin-bottom:1.2rem; }
        .highlight  { background:#f0f4fa; border-left:3px solid #1a3a5c;
                      padding:.5rem .8rem; border-radius:4px; font-size:.85rem; }
        </style>
        <div class="main-title">🔥 도시가스 도매요금 정산 계산기</div>
        <div class="sub-title">검침주기 보정 raw data 기반 용도별 구성비 &amp; 공급량 자동 산출</div>
    """, unsafe_allow_html=True)

    with st.spinner("GitHub에서 raw data 파일 목록 로딩 중…"):
        xlsx_files = get_github_file_list()

    with st.expander("🛠️ 시스템 상태 (문제 발생 시 확인)", expanded=not bool(xlsx_files)):
        status = st.session_state.get("api_status", "N/A")
        st.write(f"GitHub API 상태: `{status}` (200이면 정상)")
        if status != 200:
            st.error(f"오류: {st.session_state.get('api_error_msg', '')}")
        st.write("발견된 xlsx 파일:", xlsx_files)

    if not xlsx_files:
        st.warning("xlsx 파일을 찾지 못했습니다.")
        return

    st.markdown("---")
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("**① Raw Data 파일 선택**")
        st.caption("'전체' 또는 메인 데이터가 포함된 파일을 선택하세요 (일반용/취사용 등 부분 파일은 제외)")
        selected_file = st.selectbox("파일", xlsx_files)

    with col2:
        st.markdown("**② 기준 월 선택**")
        target_month = st.selectbox("정산 대상 월 (M컬럼)", [2, 3, 4, 5], index=1,
                                     format_func=lambda x: f"{x}월 보정공급량")

    st.markdown("**③ 수급량 입력 (GJ)**")
    supply_gj = st.number_input(
        "월 수급량 (GJ)", min_value=0.0, value=4_662_707.166,
        step=1000.0, format="%.3f",
        help="가스공사 공급량 일일보고서의 월 수급량(GJ)"
    )

    st.markdown("")

    if st.button("🔢 구성비 & 공급량 산출", type="primary", use_container_width=True):
        with st.spinner("엑셀 파일 다운로드 & 계산 중…"):
            xlsx_bytes = download_xlsx(selected_file)
            if not xlsx_bytes:
                st.error(f"파일 다운로드 실패: {selected_file}")
                return

            df_raw = read_verification_sheet(xlsx_bytes)
            if df_raw is None:
                st.error("'검증' 시트를 찾거나 파싱할 수 없습니다. 파일 구조를 확인해 주세요.")
                return

            month_col = f"M{target_month}"
            if month_col not in df_raw.columns:
                st.error(f"{target_month}월 보정공급량 컬럼이 없습니다. "
                          f"사용 가능 컬럼: {[c for c in df_raw.columns if c.startswith('M')]}")
                return

            usage_totals = allocate_by_usage(df_raw, month_col)

        if not usage_totals or sum(usage_totals.values()) == 0:
            st.error("매핑된 데이터가 없습니다. SERVICE_MAP과 raw data 구조를 확인해 주세요.")
            return

        df = calc_settlement(supply_gj, usage_totals)

        st.markdown("### 📊 용도별 공급량 배분 결과")

        def style_row(row):
            if row["용도"] in ("합계",) or row["용도"] in SUBTOTAL_LABELS:
                return ["background-color:#f0f4fa; font-weight:bold"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(style_row, axis=1).format({
                "보정공급량(원본단위)": "{:,.0f}",
                "구성비(%)":      lambda x: f"{x:.4f}%" if pd.notna(x) else "-",
                "배분공급량(GJ)": lambda x: f"{x:,.3f}" if pd.notna(x) else "-",
            }),
            use_container_width=True, height=460
        )

        total_row = df[df["용도"] == "합계"].iloc[0]
        c1, c2 = st.columns(2)
        c1.metric("입력 수급량 (GJ)", f"{supply_gj:,.3f}")
        c2.metric("배분 합계 (GJ)", f"{total_row['배분공급량(GJ)']:,.3f}")

        csv = df.to_csv(index=False, encoding="utf-8-sig")
        st.download_button(
            label="⬇️ 정산 결과 CSV 다운로드",
            data=csv.encode("utf-8-sig"),
            file_name=f"도매정산_{target_month}월.csv",
            mime="text/csv",
        )

        with st.expander("📋 원본 raw data 상세 (검증용)"):
            st.dataframe(df_raw, use_container_width=True)

    with st.expander("ℹ️ 로직 설명 & 필요 Raw Data 안내"):
        st.markdown("""
### 계산 로직
1. Raw data의 `상품` + `서비스` 조합별 N월 **보정공급량**(검침주기 일할계산 반영값) 추출
2. 가스공사 9개 용도 체계로 (상품,서비스) → 용도 매핑 후 합산
3. 합산값 ÷ 전체합계 × 100 = 구성비(%)
4. 구성비 × 수급량(GJ) = 배분 공급량(GJ)

### 용도 매핑 표
| 가스공사 용도 | 포함 (상품, 서비스) |
|---|---|
| 주택용 | 개별난방용/난방용/취사난방용(개별난방·취사), 냉난방용(주택)/중앙난방, 자가열전용, 취사용 |
| 업무난방용 | 업무난방용 |
| 일반용 | 일반용(1)/(2) 기타·동절기 |
| 냉난방공조용 | 냉난방용(업무) 기타·동절기 |
| 산업용 | 산업용 기타·동절기 |
| 수송용 | 수송용(CNG), 수송용(외주) ※BIO 제외 |
| 열병합용 | 열병합용 기타·동절기 |
| 연료전지용 | 연료전지 기타·동절기 |
| 열전용설비용 | 열전용설비용 |
| 주한미군 | 주한미군 |

### 필요한 Raw Data
**이제 판매량정산서(빌링팀)는 필요 없습니다.** 재무팀이 제공하는 `_공급량_판매량_비율자료_*.xlsx`의
**`N월검증` 시트** 하나만 있으면 가스공사 정산서를 100% 가까운 정확도로 재현할 수 있습니다 (검증 오차 0.2%p 이내).

다만 다음 사항을 매월 확인하세요:
- 파일에 `'전체'`가 포함된 메인 파일을 사용 (일반용/취사용만 있는 부분 파일 제외)
- 시트명이 `N월검증` 형식인지 확인 (자동 탐지하지만 이름이 크게 다르면 실패할 수 있음)
- `상품`/`서비스` 조합이 기존과 다르면 `SERVICE_MAP` 갱신 필요
        """)


if __name__ == "__main__":
    main()
