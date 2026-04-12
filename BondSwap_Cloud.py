import streamlit as st
import pandas as pd
import numpy as np

# 페이지 설정
st.set_page_config(layout="wide", page_title="포지션 현황")

# --- 비밀번호 확인 기능 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if st.session_state["password_correct"]:
        return True

    st.title("보안 인증")
    password = st.text_input("비밀번호를 입력하세요.", type="password")
    if st.button("로그인"):
        if password == "123456":
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

# --- 완벽한 병합 및 스타일을 위한 HTML 렌더링 함수 ---
def render_merged_html(df):
    # 인덱스 이름(큰분류, 세부분류 등)을 헤더에서 깔끔하게 제거
    df.index.names = [None] * len(df.index.names)
    
    # Pandas의 to_html은 MultiIndex를 자동으로 병합(rowspan) 해줍니다.
    raw_html = df.to_html(float_format=lambda x: f"{x:,.0f}")
    
    # 세부분류 구분이 확 띄도록 커스텀 CSS 적용
    custom_css = """
    <style>
        .custom-table {
            width: 100%;
            border-collapse: collapse;
            font-family: 'Arial', sans-serif;
            font-size: 14px;
        }
        .custom-table th, .custom-table td {
            border: 1px solid #777; /* 기본 얇은 테두리 */
            padding: 8px;
            text-align: center !important;
            vertical-align: middle !important; /* 병합된 셀 텍스트를 정중앙에 위치 */
        }
        .custom-table thead th {
            background-color: #d5d8dc;
            border: 2px solid #222; /* 열 헤더 진한 테두리 */
            font-weight: bold;
        }
        .custom-table tbody th {
            background-color: #eaeded;
            border: 2px solid #222; /* 행 분류(병합된 셀) 진한 테두리 */
            font-weight: bold;
        }
        /* 숫자가 들어가는 일반 데이터 셀은 흰색 배경 */
        .custom-table tbody td {
            background-color: #ffffff;
        }
    </style>
    """
    
    # Pandas가 생성한 테이블에 custom-table 클래스 부여
    styled_html = raw_html.replace('<table border="1" class="dataframe">', '<table class="custom-table">')
    return custom_css + styled_html


if check_password():
    st.title("원화채권 포지션 현황")

    file_path = "포지션 보고양식 예시.xlsx"

    @st.cache_data
    def load_and_preprocess_data():
        sheets = ['9999', '9994', '9992', '9988', '7120(원화)']
        df_list = []
        for sheet in sheets:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet)
                df_list.append(df)
            except Exception as e:
                st.warning(f"'{sheet}' 시트를 불러오지 못했습니다. ({e})")
        
        if not df_list: return pd.DataFrame()
        data = pd.concat(df_list, ignore_index=True)
        data = data[data['펀드코드'] != 8018].copy()
        
        def classify_bond(row):
            ctype, name = str(row['채권종류']).strip(), str(row['종목명']).strip()
            if ctype in ['국채', '통안채', '지방채']: return '국고/통안채'
            elif '중앙회' in name or ctype == '특수채': return '특수채'
            elif ctype == '금융채':
                if any(x in name for x in ['중소기업은행', '기업은행', '산업금융채권', '산금']): return '특은채'
                elif any(x in name for x in ['카드', '캐피탈']): return '여전채'
                elif '은행' in name: return '시은채'
                else: return '기타금융채'
            return '기타'

        data['분류_채권종류'] = data.apply(classify_bond, axis=1)
        data['평가일자'] = pd.to_datetime(data['평가일자'])
        data['만기일자'] = pd.to_datetime(data['만기일자'])
        data['잔존년수'] = (data['만기일자'] - data['평가일자']).dt.days / 365.0
        
        def classify_maturity(years):
            bins = [0, 0.25, 0.5, 1, 1.5, 2, 2.5, 3, 5, 10, 20, 100]
            labels = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']
            for i, b in enumerate(bins[1:]):
                if years <= b: return labels[i]
            return '30Y'
            
        data['분류_잔존만기'] = data['잔존년수'].apply(classify_maturity)
        data['포지션(억원)'] = data['수량'] / 100000.0
        
        def classify_fund(code):
            if code == 8010: return 'RP운용', '대고객 RP(8010)', '대고객 RP(8010)'
            elif code == 8013: return 'RP운용', 'CMA RP(8013)', 'CMA RP(8013)'
            elif code == 9994: return 'RP운용', '기관RP(9994)', '기관RP(9994)'
            elif code == 7120: return 'RP운용', '외화RP(7120)', '외화RP(7120)'
            elif code in [8001, 8008]: return '상품채권운용', '자격', 'PD(8001)' if code==8001 else '소액(8008)'
            elif code in [8007, 8016, 8011, 8019]: return '상품채권운용', '일반 Prop', \
                {8007:'팀운용', 8016:'부서공통', 8011:'Prop1', 8019:'Prop2'}[code]
            return '기타', '기타', '기타'

        data[['큰분류', '세부분류', '세세부']] = data.apply(lambda row: pd.Series(classify_fund(row['펀드코드'])), axis=1)
        return data

    df = load_and_preprocess_data()

    if not df.empty:
        maturity_order = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']
        rp_order = ['대고객 RP(8010)', 'CMA RP(8013)', '기관RP(9994)', '외화RP(7120)']
        
        tab1, tab2 = st.tabs(["📊 RP운용 현황", "📈 상품채권운용 현황"])

        with tab1:
            rp_df = df[df['큰분류'] == 'RP운용'].copy()
            if not rp_df.empty:
                rp_df['세부분류'] = pd.Categorical(rp_df['세부분류'], categories=rp_order, ordered=True)
                rp_pivot = pd.pivot_table(rp_df, values='포지션(억원)', 
                                         index=['큰분류', '세부분류', '분류_채권종류'], 
                                         columns='분류_잔존만기', aggfunc='sum', fill_value=0).sort_index(level='세부분류')
                
                cols = [c for c in maturity_order if c in rp_pivot.columns]
                rp_pivot = rp_pivot[cols]
                rp_pivot['합계'] = rp_pivot.sum(axis=1)
                
                # st.dataframe 대신 HTML 렌더링 사용
                html_table = render_merged_html(rp_pivot)
                st.markdown(html_table, unsafe_allow_html=True)
            else:
                st.info("데이터가 없습니다.")

        with tab2:
            prop_df = df[df['큰분류'] == '상품채권운용']
            if not prop_df.empty:
                prop_pivot = pd.pivot_table(prop_df, values='포지션(억원)', 
                                           index=['큰분류', '세부분류', '세세부', '분류_채권종류'], 
                                           columns='분류_잔존만기', aggfunc='sum', fill_value=0)
                
                cols = [c for c in maturity_order if c in prop_pivot.columns]
                prop_pivot = prop_pivot[cols]
                prop_pivot['합계'] = prop_pivot.sum(axis=1)
                
                # st.dataframe 대신 HTML 렌더링 사용
                html_table = render_merged_html(prop_pivot)
                st.markdown(html_table, unsafe_allow_html=True)
            else:
                st.info("데이터가 없습니다.")
    else:
        st.error("데이터 로드 실패")
