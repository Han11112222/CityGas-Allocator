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
        ("냉난방용(주택)", "냉난방(하절기)"),
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
        ("냉난방용(업무)", "냉난방(하절기)"),
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
def read_verification_sheet(xlsx_bytes: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
    """
    원본 raw data 시트를 읽어 (상품, 서비스, N월보정공급량...) DataFrame 반환.
    시트명이 'N월검증' 형식으로 매월 바뀌므로 '검증'이 포함된 시트를 자동 탐색한다.
    헤더 행을 자동 탐지하여 '상품'/'서비스' 컬럼을 찾는다.
    일부 파일은 '상품/용도/서비스' 3컬럼 구조(부분 서브셋)이므로 이 경우도 흡수한다.
    """
    try:
        xl = pd.ExcelFile(BytesIO(xlsx_bytes))
    except Exception as e:
        st.error(f"엑셀 파일 열기 실패: {e}")
        return None

    if sheet_name is None or sheet_name not in xl.sheet_names:
        candidates = [s for s in xl.sheet_names if "검증" in s]
        if not candidates:
            return None
        sheet_name = candidates[0]

    raw = xl.parse(sheet_name, header=None)

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

    month_cols = {}
    for col in data_rows.columns:
        col_str = str(col)
        m = re.search(r"(\d{1,2})월보정공급량", col_str)
        if m and "합계" in col_str:
            month_cols[int(m.group(1))] = col

    has_yongdo = "용도" in data_rows.columns

    data_rows["상품"] = data_rows["상품"].ffill()
    if has_yongdo:
        data_rows["용도"] = data_rows["용도"].ffill()

    base_cols = ["상품", "서비스"] + (["용도"] if has_yongdo else [])
    result_cols = base_cols + list(month_cols.values())
    df = data_rows[result_cols].copy()
    df = df.dropna(subset=["서비스"])
    df = df[df["상품"] != "총합계"]

    rename_map = {v: f"M{k}" for k, v in month_cols.items()}
    df = df.rename(columns=rename_map)

    for c in df.columns:
        if c.startswith("M"):
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    return df


def is_full_dataset(df: pd.DataFrame) -> bool:
    """파싱된 raw data가 전체 상품 종류를 충분히 포함하는지 확인.
       서브셋 파일(일반용만, 취사용만 등)을 거르기 위한 안전장치."""
    if df is None or "상품" not in df.columns:
        return False
    return df["상품"].nunique() >= 10


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


def apply_prior_month_correction(
    usage_totals: dict,
    prior_actual_ratio: dict,
    prior_raw_totals: dict,
) -> dict:
    """
    전월 raw값과 전월 가스공사 확정 구성비(실제값)를 비교해 용도별 '환산배율'을 역산한 뒤,
    이번 달 raw값을 그 배율로 나눠 보정한다.

        전월배율(용도) = 전월raw(용도) / 전월실제MJ(용도)
        보정값(용도)   = 당월raw(용도) / 전월배율(용도)

    raw data 자체가 추정치이므로 매월 ±2~5% 수준의 흔들림이 있는데,
    이 흔들림이 용도별로 비슷한 패턴(공통 시간추세)을 보이는 경향이 있어
    전월 오차를 그대로 당월에 역보정하면 정확도가 개선되는 경우가 많다(검증 결과 약 20% 오차 감소).
    """
    prior_total_raw = sum(prior_raw_totals.values())
    prior_total_ratio = sum(prior_actual_ratio.values())
    if prior_total_raw == 0 or prior_total_ratio == 0:
        return usage_totals

    corrected = {}
    for label, raw_val in usage_totals.items():
        prior_raw = prior_raw_totals.get(label, 0)
        prior_ratio = prior_actual_ratio.get(label, 0)
        if prior_raw > 0 and prior_ratio > 0:
            prior_actual_mj = prior_total_raw * (prior_ratio / prior_total_ratio)
            factor = prior_raw / prior_actual_mj if prior_actual_mj else 1.0
            corrected[label] = raw_val / factor if factor > 0 else raw_val
        else:
            corrected[label] = raw_val
    return corrected


# ──────────────────────────────────────────
# 정산 계산
# ──────────────────────────────────────────
def calc_settlement(supply_gj: float, usage_totals: dict) -> pd.DataFrame:
    """
    usage_totals: {용도: raw 보정공급량(MJ, 일할계산 반영 전체 판매량 기준)}
    구성비는 raw 값들 간의 비율로 산출하므로, raw 값에 공통적으로 곱해지는
    '증량 계수'(전월 검침→당월 검침 사이 부피 증가분)는 구성비 계산 시 자동 약분되어
    영향을 주지 않는다. 즉 증량의 절대크기를 몰라도 구성비 산출에는 지장이 없다.
    """
    grand_total = sum(usage_totals.values())
    rows = []
    for label in TARGET_ORDER:
        val = usage_totals.get(label, 0)
        ratio = (val / grand_total * 100) if grand_total > 0 else None
        alloc_gj = supply_gj * ratio / 100 if ratio is not None else None
        # 참고용: raw값을 GJ로 단순환산(÷1000)했을 때 수급량 배분 GJ와의 배율
        # → 이 배율이 용도마다 비슷하면 raw가 이미 "증량 포함 판매량"이라는 방증
        implied_factor = (val / 1000) / alloc_gj if alloc_gj else None
        rows.append({
            "용도":            label,
            "원본판매량(MJ)":  round(val, 0),
            "구성비(%)":       round(ratio, 4) if ratio is not None else None,
            "배분공급량(GJ)":  round(alloc_gj, 3) if alloc_gj is not None else None,
            "환산배율(참고)":  round(implied_factor, 3) if implied_factor else None,
        })
    total_row = {
        "용도": "합계",
        "원본판매량(MJ)": round(grand_total, 0),
        "구성비(%)": 100.0 if grand_total > 0 else None,
        "배분공급량(GJ)": round(supply_gj, 3) if grand_total > 0 else None,
        "환산배율(참고)": round((grand_total/1000)/supply_gj, 3) if supply_gj else None,
    }
    rows.append(total_row)
    return pd.DataFrame(rows)


# ──────────────────────────────────────────
# UI
# ──────────────────────────────────────────
def main():
    st.set_page_config(page_title="DSE 수급량 용도별 배분", page_icon="🔥", layout="wide")

    st.markdown("""
        <style>
        .main-title { font-size:2.4rem; font-weight:800; color:#1a3a5c;
                      border-bottom:3px solid #e65c00; padding-bottom:.4rem; margin-bottom:.3rem; }
        .sub-title  { font-size:1rem; color:#666; margin-bottom:1.4rem; }
        .step-title { font-size:1.25rem; font-weight:700; color:#1a3a5c; margin-bottom:.4rem; }
        .highlight  { background:#f0f4fa; border-left:3px solid #1a3a5c;
                      padding:.5rem .8rem; border-radius:4px; font-size:.85rem; }
        </style>
        <div class="main-title">🔥 DSE 수급량 용도별 배분</div>
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

    # 메인(전체) 파일만 후보로 사용 — 서브셋(일반용/취사용 등) 파일은 자동 제외
    SUBSET_MARKERS = ["일반용", "취사용", "취사추정"]
    main_candidates = [
        f for f in xlsx_files
        if "전체" in f and not any(marker in f for marker in SUBSET_MARKERS)
    ]
    if not main_candidates:
        main_candidates = [f for f in xlsx_files if "전체" in f] or xlsx_files

    # 파일명에서 "N월" 추출 → {월: 파일경로} 매핑 (동일 월에 여러 파일이 있으면 가장 마지막 파일 우선)
    month_to_file: dict[int, str] = {}
    for f in sorted(main_candidates):
        m = re.search(r"(\d{1,2})월", f)
        if m:
            month_to_file[int(m.group(1))] = f

    st.markdown('<div class="step-title">① 기준 월 선택</div>', unsafe_allow_html=True)
    available_months = sorted(month_to_file.keys()) if month_to_file else [2, 3, 4, 5]
    default_idx = len(available_months) - 1 if available_months else 0
    target_month = st.selectbox(
        "정산 대상 월", available_months, index=default_idx,
        format_func=lambda x: f"{x}월 보정공급량", label_visibility="collapsed",
    )

    # 기준 월에 맞는 raw data 파일 자동 선택
    selected_file = month_to_file.get(target_month)
    if selected_file:
        st.markdown(
            f'<div class="highlight">📁 자동 선택된 Raw Data 파일: <b>{selected_file}</b></div>',
            unsafe_allow_html=True,
        )
    else:
        st.warning(f"{target_month}월에 해당하는 '전체' raw data 파일을 자동으로 찾지 못했습니다. "
                    "직접 선택해 주세요.")
        selected_file = st.selectbox("파일 직접 선택", sorted(xlsx_files))

    st.markdown("")
    st.markdown('<div class="step-title">② 수급량 입력 (GJ)</div>', unsafe_allow_html=True)
    supply_gj = st.number_input(
        "월 수급량 (GJ)", min_value=0.0, value=4_662_707.166,
        step=1000.0, format="%.3f", label_visibility="collapsed",
        help="가스공사 공급량 일일보고서의 월 수급량(GJ)"
    )

    st.markdown("")
    st.markdown('<div class="step-title">③ 정밀 보정 (선택)</div>', unsafe_allow_html=True)
    use_correction = st.toggle(
        "전월 확정 구성비로 보정하기",
        value=False,
        help="raw data(보정공급량)는 추정치라 매월 ±2~5% 수준의 흔들림이 있습니다. "
             "직전월 가스공사 확정 구성비를 입력하면, 그 오차 패턴을 역산해 "
             "이번 달 구성비를 보정합니다 (검증 결과 오차 약 20% 감소)."
    )

    prior_ratio_input = {}
    prior_raw_totals = {}
    if use_correction:
        prev_month = target_month - 1 if target_month > 1 else 12
        prev_file = month_to_file.get(prev_month)
        if not prev_file:
            st.warning(f"{prev_month}월 raw data 파일을 자동으로 찾지 못해 보정을 적용할 수 없습니다.")
            use_correction = False
        else:
            st.caption(f"📁 전월({prev_month}월) raw data: {prev_file}")
            st.caption("전월 가스공사 확정 정산서의 '구성비(%)' 값을 그대로 입력하세요.")
            with st.expander(f"{prev_month}월 확정 구성비 입력", expanded=True):
                cols = st.columns(5)
                for i, label in enumerate(TARGET_ORDER):
                    with cols[i % 5]:
                        prior_ratio_input[label] = st.number_input(
                            label, min_value=0.0, max_value=100.0, value=0.0,
                            step=0.0001, format="%.4f", key=f"prior_{label}"
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

            if not is_full_dataset(df_raw):
                st.error(
                    f"⚠️ 이 파일은 전체 데이터가 아닌 것으로 보입니다 "
                    f"(상품 종류 {df_raw['상품'].nunique()}개만 발견됨, 전체는 보통 15개 이상). "
                    f"'전체'가 포함된 메인 파일을 다시 선택해 주세요."
                )
                with st.expander("발견된 상품 목록 보기"):
                    st.write(sorted(df_raw["상품"].dropna().unique().tolist()))
                return

            month_col = f"M{target_month}"
            if month_col not in df_raw.columns:
                st.error(f"{target_month}월 보정공급량 컬럼이 없습니다. "
                          f"사용 가능 컬럼: {[c for c in df_raw.columns if c.startswith('M')]}")
                return

            usage_totals = allocate_by_usage(df_raw, month_col)

            applied_correction = False
            if use_correction and prior_ratio_input and sum(prior_ratio_input.values()) > 0:
                prev_month = target_month - 1 if target_month > 1 else 12
                prev_file = month_to_file.get(prev_month)
                prev_bytes = download_xlsx(prev_file)
                if prev_bytes:
                    df_prev = read_verification_sheet(prev_bytes)
                    prev_col = f"M{prev_month}"
                    if df_prev is not None and prev_col in df_prev.columns:
                        prior_raw_totals = allocate_by_usage(df_prev, prev_col)
                        usage_totals = apply_prior_month_correction(
                            usage_totals, prior_ratio_input, prior_raw_totals
                        )
                        applied_correction = True

        if not usage_totals or sum(usage_totals.values()) == 0:
            st.error("매핑된 데이터가 없습니다. SERVICE_MAP과 raw data 구조를 확인해 주세요.")
            return

        if applied_correction:
            st.success("✅ 전월 확정 구성비 기반 보정이 적용되었습니다.")

        df = calc_settlement(supply_gj, usage_totals)

        st.markdown("### 📊 용도별 공급량 배분 결과")

        def style_row(row):
            if row["용도"] == "합계":
                return ["background-color:#f0f4fa; font-weight:bold"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df.style.apply(style_row, axis=1).format({
                "원본판매량(MJ)": "{:,.0f}",
                "구성비(%)":      lambda x: f"{x:.4f}%" if pd.notna(x) else "-",
                "배분공급량(GJ)": lambda x: f"{x:,.3f}" if pd.notna(x) else "-",
                "환산배율(참고)": lambda x: f"{x:.3f}" if pd.notna(x) else "-",
            }),
            use_container_width=True, height=460, hide_index=True
        )
        st.caption(
            "💡 **환산배율(참고)**: 원본판매량(MJ)÷1000을 배분공급량(GJ)으로 나눈 값입니다. "
            "용도별로 이 값이 거의 비슷하다면, 검침주기 사이 부피증가(증량)가 "
            "용도와 무관하게 비슷한 비율로 반영되었다고 추정할 수 있습니다 — "
            "다만 정확한 증량 적용 기준은 재무팀만 알고 있어 추정치입니다."
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
            st.dataframe(df_raw, use_container_width=True, hide_index=True)

    with st.expander("ℹ️ 로직 설명 & 증량(增量) 처리 방식"):
        st.markdown("""
### 계산 로직
1. Raw data의 `상품` + `서비스` 조합별 **N월 보정공급량**(검침주기 일할계산 반영, 증량 포함 추정 판매량) 추출
2. 가스공사 9개 용도 체계로 (상품,서비스) → 용도 매핑 후 합산
3. 합산값 ÷ 전체합계 × 100 = **구성비(%)**
4. 구성비 × 수급량(GJ, 고정값) = **배분 공급량(GJ)**

### 증량(增量)을 모르는데 왜 구성비 계산이 가능한가?
재무팀 설명에 따르면 수급량→판매량 환산 시 "증량"(검침주기 사이 부피 증가분)이 +α로 더해집니다.
이 증량의 절대 크기는 공유받지 못했지만, **구성비를 계산할 때는 분자·분모에서 증량 배율이
약분되어 사라지므로 몰라도 무방**합니다.

```
구성비 = (용도 A의 raw판매량) / (전체 raw판매량 합계)
       = (용도 A의 실판매량 × 증량배율) / (전체 실판매량 × 증량배율)
       = (용도 A의 실판매량) / (전체 실판매량)   ← 증량배율 약분됨
```

단, 이 약분이 성립하려면 **증량배율이 모든 용도에 동일하게 적용**되어야 합니다. 실제 4개월
데이터로 역산한 결과, 열전용설비용을 제외한 9개 용도는 "환산배율"이 매월 비슷한 범위
(예: 3월 기준 968~1047)에 분포해 **거의 동일한 배율로 증량이 적용된다는 정황**이 확인됩니다.
다만 **열전용설비용만 매월 환산배율이 크게 이탈**하는데, 이는 raw data 자체에 해당 용도의
물량이 일부 누락되었을 가능성을 시사합니다 (재무팀 확인 필요).

### M-1, M월 2개월만 쓴다는 재무팀 설명에 대해
재무팀이 "M-1, M월 2개 검침 데이터만 사용한다"고 했지만, 본 raw data의 `N월보정공급량`
컬럼들은 **이미 일할계산으로 당월 귀속분이 분리·완성된 값**입니다. 실제로 (전월컬럼+당월컬럼)을
단순 평균하면 검증 오차가 0.5%p → 26%p로 폭증하는 것을 확인했습니다. 따라서 **"당월(M) 컬럼
하나만 사용"하는 현재 방식이 재무팀의 M-1/M 로직이 이미 내재적으로 반영된 최종 결과값**이라고
판단됩니다 (재무팀이 말하는 M-1/M 처리는 이 컬럼을 만드는 **이전 단계**에서 이미 끝난 것으로 추정).

### 용도 매핑 표
| 가스공사 용도 | 포함 (상품, 서비스) |
|---|---|
| 주택용 | 개별난방용/난방용/취사난방용(개별난방·취사), 냉난방용(주택)/중앙난방·하절기, 자가열전용, 취사용 |
| 업무난방용 | 업무난방용 |
| 일반용 | 일반용(1)/(2) 기타·동절기 |
| 냉난방공조용 | 냉난방용(업무) 기타·동절기·하절기 |
| 산업용 | 산업용 기타·동절기 |
| 수송용 | 수송용(CNG), 수송용(외주) ※BIO 제외 |
| 열병합용 | 열병합용 기타·동절기 |
| 연료전지용 | 연료전지 기타·동절기 |
| 열전용설비용 | 열전용설비용 |
| 주한미군 | 주한미군 |

### 필요한 Raw Data
**판매량정산서(빌링팀)는 더 이상 필요 없습니다.** 재무팀이 제공하는 `_공급량_판매량_비율자료_*.xlsx`의
**`N월검증` 시트** 하나만 있으면 가스공사 정산서를 평균 0.2%p 이내 오차로 재현할 수 있습니다.

매월 확인할 사항:
- 파일에 `'전체'`가 포함된 메인 파일을 사용 (일반용/취사용만 있는 부분 파일은 제외)
- `상품`/`서비스` 조합에 새 이름이 생기면(예: '냉난방(하절기)') `SERVICE_MAP` 갱신 필요
- 열전용설비용 오차가 계속 크면 재무팀에 raw data 누락 여부 확인 권장
        """)


if __name__ == "__main__":
    main()
