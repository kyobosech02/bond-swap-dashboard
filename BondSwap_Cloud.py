import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime

# 1. 페이지 설정 (모바일 대응을 위한 layout 설정)
st.set_page_config(page_title="K-Bond Mobile Analysis", page_icon="📈", layout="wide")

# 모바일 가독성을 위한 커스텀 CSS
st.markdown("""
    <style>
    /* 메트릭 카드 폰트 조절 */
    [data-testid="stMetricValue"] { font-size: 1.5rem !important; }
    /* 모바일에서 표의 폰트 크기 및 여백 최적화 */
    .dataframe { font-size: 12px !important; }
    /* 탭 메뉴 폰트 크기 조절 */
    button[data-baseweb="tab"] { font-size: 14px !important; padding: 10px 5px !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 0. 비밀번호 인증 로직 (기존 유지)
# ==========================================
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    
    if st.session_state["password_correct"]:
        return True

    def password_entered():
        if st.session_state["password_input"] == "kyoboh02":
            st.session_state["password_correct"] = True
            del st.session_state["password_input"]
        else:
            st.error("😕 비밀번호가 틀렸습니다.")

    st.title("🔒 Access Required")
    st.text_input("비밀번호를 입력하세요", type="password", key="password_input", on_change=password_entered)
    return False

# ==========================================
# 1. 데이터 로드 (캐싱 및 오류 처리)
# ==========================================
@st.cache_data
def load_data():
    # 실제 환경에서는 file_path = "Data.xlsx"를 사용하세요.
    # 여기서는 구조 이해를 위해 임시 로직만 유지합니다.
    try:
        file_path = "Data.xlsx"
        # ... (기존 load_data 로직 수행) ...
        # 샘플 반환 (실제 코드 연결 필요)
        return bond_dict, irs_dict
    except Exception as e:
        st.error(f"파일을 읽을 수 없습니다: {e}")
        return None, None

# ==========================================
# 메인 대시보드
# ==========================================
if check_password():
    # 데이터 로드 (실제 데이터 딕셔너리가 있다고 가정)
    # bond_data, irs_data = load_data()
    
    # [임시 설정 - 실제 데이터 연동 시 삭제]
    bond_list = ['국고채권', '은행채 AAA', '공사채 AAA']
    maturities_list = ['3M', '6M', '9M', '1Y', '1.5Y', '2Y', '3Y', '4Y', '5Y']
    x_numeric = [0.25, 0.5, 0.75, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]

    # 사이드바 (모바일에서는 왼쪽 상단 햄버거 메뉴로 숨겨짐)
    with st.sidebar:
        st.header("⚙️ 분석 옵션")
        st.caption("교보증권 채권운용부 유지민")
        
        with st.expander("1) Bond-Swap 분석 설정", expanded=True):
            selected_bond = st.selectbox("채권 종류", bond_list)
            selected_mat = st.selectbox("분석 만기", maturities_list, index=3)
            # 날짜 선택기를 간소화 (한 줄에 배치)
            d_col1, d_col2 = st.columns(2)
            with d_col1:
                start_date = st.date_input("시작일", datetime.date.today() - datetime.timedelta(days=365))
            with d_col2:
                end_date = st.date_input("종료일", datetime.date.today())

    # 메인 타이틀
    st.title("📈 Bond-Swap Dash")

    # 탭 메뉴 (모바일 터치에 최적화)
    tab1, tab2 = st.tabs(["📊 Bond-Swap", "💳 Credit Spread"])

    with tab1:
        # --- 핵심 지표 (Metrics) : 모바일에서 가장 먼저 보임 ---
        st.subheader(f"📍 {selected_mat} 실시간 요약")
        m1, m2, m3 = st.columns(3)
        # 실제 데이터 연동 (final_df에서 추출)
        m1.metric("Spread", "15.2 bp", "1.2 bp")
        m2.metric("채권 금리", "3.450 %", "-0.005 %")
        m3.metric("IRS 금리", "3.602 %", "0.007 %")

        # --- 차트 영역 ---
        # 모바일에서는 차트를 하나씩 세로로 배치하는 것이 좋음
        st.divider()
        st.subheader("📉 금리 및 스프레드 추이")
        
        # fig1 설정 (모바일 최적화 레이아웃)
        fig1 = make_subplots(specs=[[{"secondary_y": True}]])
        # ... (기존 fig1 Scatter 추가 로직) ...
        # 예시 데이터 추가
        fig1.add_trace(go.Scatter(name="채권", y=[3.4, 3.45, 3.42]), secondary_y=False)
        fig1.add_trace(go.Scatter(name="Spread", y=[12, 15, 14], fill='tozeroy'), secondary_y=True)

        fig1.update_layout(
            height=350, # 모바일에서 한 화면에 차트와 설명이 들어오도록 높이 조절
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode="x unified"
        )
        st.plotly_chart(fig1, use_container_width=True, config={'displayModeBar': False})

        # --- 상세 데이터 (Expander로 숨김 처리) ---
        with st.expander("📄 상세 데이터 및 통계 확인"):
            # 최근 10일치만 우선 보여주기 등 최적화 가능
            st.write("최근 데이터 리스트")
            # st.dataframe(styled_df) # 기존 스타일 적용 데이터프레임

    with tab2:
        st.subheader("💳 Credit Spread 분석")
        
        # 모바일용 만기 선택 필터 (표가 너무 길어질 경우 대비)
        view_mats = st.multiselect("확인할 만기 선택", maturities_list, default=['1Y', '2Y', '3Y', '5Y'])
        
        # --- 핵심 스프레드 요약 표 ---
        # 모바일에서는 가로로 긴 표보다, 핵심 정보만 추린 카드 형태나 좁은 표가 유리함
        # 여기서는 기존 표를 유지하되 스타일로 가독성 확보
        # ... (기존 t2_final_df 데이터 처리 로직) ...
        
        st.caption("💡 만기별 Spread 현황 (bp)")
        # 예시용 간소화된 데이터프레임
        summary_data = {
            "만기": view_mats,
            "현재": [15.1, 18.2, 20.5, 22.1],
            "전일비": ["▲ 0.5", "▼ 1.2", "-", "▲ 0.8"],
            "Percentile": ["80%", "45%", "60%", "90%"]
        }
        st.table(pd.DataFrame(summary_data).set_index("만기"))

        # --- 만기별 커브 차트 ---
        st.subheader("📉 만기별 Spread 커브")
        fig2 = go.Figure()
        # ... (기존 fig2 로직) ...
        fig2.update_layout(
            height=350,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})

    # 하단 정보 (고정형 캡션)
    st.markdown("---")
    st.caption("© 2026 Kyobo Securities Bond Management Dept. | PC/Mobile Optimized")
