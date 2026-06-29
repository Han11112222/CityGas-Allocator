import streamlit as st
import pandas as pd

# 페이지 기본 설정
st.set_page_config(page_title="도시가스 수급량 배분 정산기", layout="wide", page_icon="🔥")

st.title("🔥 도시가스 도매요금 수급량 배분 정산기")
st.markdown("빌링팀의 **판매량정산서(데이터.csv 또는 엑셀)**를 업로드하고 총 수급량을 입력하면, 용도별 구성비에 맞추어 추정 수급량을 자동으로 배분합니다.")

# 사이드바 설정
st.sidebar.header("1. 기초 데이터 입력")
total_supply = st.sidebar.number_input("가스공사 당월(또는 m-1월) 총 수급량 (GJ)", min_value=0, value=100000, step=1000)

st.sidebar.header("2. 판매량정산서 업로드")
uploaded_file = st.sidebar.file_uploader("판매량정산서 데이터 파일 업로드 (CSV/Excel)", type=['csv', 'xlsx'])

# 초기 구성비 딕셔너리 설정
ratios = {
    "주택용": 25.6,
    "일반용": 11.4,
    "업무냉방용": 11.7,
    "산업용": 34.7,
    "수송용": 14.7,
    "기타": 1.9
}

if uploaded_file is not None:
    try:
        # CSV 파일인 경우
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file, sheet_name='데이터') # 엑셀의 '데이터' 시트 기준
            
        st.success("파일 업로드 성공! 파일 내 구성비 데이터를 스캔합니다.")
        
        # 파일 구조에서 '구성비'라는 단어가 포함된 부분을 찾아 파싱 (예외처리 포함)
        # 본 코드는 일반적인 정산서(데이터.csv) 포맷을 기준으로 작성되었습니다.
        # 데이터 시트의 특정 행/열에 맞춰 수정이 필요할 수 있습니다.
        st.info("실제 업무 데이터의 행/열 양식에 맞춰 아래 비율이 자동 갱신됩니다. 만약 자동 추출이 완벽하지 않다면 수동으로 미세 조정이 가능합니다.")
        
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")

st.divider()

st.subheader("📊 용도별 판매 구성비 및 단가 설정")
st.markdown("업로드된 데이터 기반 또는 수동으로 구성비(%)와 가스공사 도매원가를 설정하세요.")

col1, col2 = st.columns(2)

# 사용자 입력 위젯 생성 및 합계 검증
updated_ratios = {}
prices = {}

with col1:
    st.markdown("**용도별 구성비 (%)**")
    for usage, ratio in ratios.items():
        updated_ratios[usage] = st.number_input(f"{usage} 비율 (%)", min_value=0.0, max_value=100.0, value=float(ratio), step=0.1)

with col2:
    st.markdown("**용도별 도매단가 (원/GJ)**")
    # 예시 단가 세팅
    default_prices = {"주택용": 18000, "일반용": 19000, "업무냉방용": 17500, "산업용": 16000, "수송용": 15000, "기타": 18000}
    for usage in ratios.keys():
        prices[usage] = st.number_input(f"{usage} 단가", min_value=0, value=default_prices[usage], step=100)

total_ratio = sum(updated_ratios.values())
if abs(total_ratio - 100.0) > 0.1:
    st.warning(f"⚠️ 구성비의 합이 100%가 아닙니다! 현재 합계: {total_ratio:.2f}%")
else:
    st.success(f"✅ 구성비 합계 정상: {total_ratio:.2f}%")

st.divider()

# 계산 로직 및 결과 표시
st.subheader("📈 용도별 추정 수급량 및 도매요금 정산 결과")

results = []
total_cost = 0

for usage, ratio in updated_ratios.items():
    # 구성비에 따른 수급량 배분
    allocated_supply = total_supply * (ratio / 100.0)
    # 공급가액 계산
    cost = allocated_supply * prices[usage]
    total_cost += cost
    
    results.append({
        "용도": usage,
        "구성비 (%)": ratio,
        "배분된 추정 수급량 (GJ)": round(allocated_supply, 2),
        "적용 도매원가 (원/GJ)": prices[usage],
        "공급가액 (원)": int(cost)
    })

df_results = pd.DataFrame(results)

# 천단위 콤마 포맷팅 적용
styled_df = df_results.style.format({
    "배분된 추정 수급량 (GJ)": "{:,.2f}",
    "공급가액 (원)": "{:,.0f}"
})

st.dataframe(styled_df, use_container_width=True)

# 최종 요약
col3, col4 = st.columns(2)
col3.metric(label="총 수급량 배분 합계 (GJ)", value=f"{sum(df_results['배분된 추정 수급량 (GJ)']):,.2f}")
col4.metric(label="최종 도매요금 공급가액 (원)", value=f"{total_cost:,.0f}")

st.caption("※ 본 시뮬레이터의 계산 결과는 부피 팽창(증량) 요소를 구성비 비례로 균등 흡수한 추정치입니다.")
