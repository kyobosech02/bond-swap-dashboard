import streamlit as st
import pandas as pd
import numpy as np
import os

# 1. 페이지 설정
st.set_page_config(layout="wide", page_title="포지션 현황 관리 시스템")

# --- 2. 비밀번호 확인 기능 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    
    if st.session_state["password_correct"]:
        return True

    st.title("🔒 보안 인증")
    password = st.text_input("비밀번호를 입력하세요.", type="password")
    if st.button("로그인"):
        if password == "123456":
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False

# --- 3. HTML 렌더링 함수 (셀 병합 및 테두리 스타일) ---
def render_merged_html(df):
    # 컬럼 헤더 상단의 인덱스 명칭 제거 (계단식 깨짐 방지)
    df.columns.name = None
    
    # Pandas의 to_html은 MultiIndex를 rowspan으로 자동 병합함
    # index_names=False가 좌측 상단 깨짐을 해결하는 핵심 옵션입니다.
    raw_html = df.to_html(float_format=lambda x: f"{x:,.0f}", index_names=False)
    
    custom_css = """
    <style>
        .merged-table {
            width: 100%;
            border-collapse: collapse;
            font-family: 'Malgun Gothic', sans-serif;
            font-size: 13px;
            margin-top: 20px;
        }
        .merged-table th, .merged-table td {
            border: 1px solid #444; /* 기본 테두리 */
            padding: 10px 5px;
            text-align: center !important;
            vertical-align: middle !important;
        }
        /* 헤더 스타일 */
        .merged-table thead th {
            background-color: #2c3e50;
            color: white;
            border: 2px solid #000;
            font-weight: bold;
        }
        /* 왼쪽 분류(병합된 셀) 스타일 */
        .merged-table tbody th {
            background-color: #f8f9fa;
            border: 2px solid #000; /* 분류 경계선을 진하게 */
            font-weight: bold;
            color: #333;
        }
        /* 데이터 셀 스타일 */
        .merged-table tbody td {
            background-color: #ffffff;
        }
        /* 합계 컬럼 강조 */
        .merged-table td:last-child, .merged-table th:last-child {
            background-color: #fef5e7;
            font-weight: bold;
        }
    </style>
    """
    
    styled_html = raw_html.replace('<table border="1" class="dataframe">', '<table class="merged-table">')
    return custom_css + styled_html

# --- 4. 메인 로직 시작 ---
if check_password():
    st.title("📊 원화채권 포지션 현황")

    # 파일 경로 설정 (실제 환경에 맞춰 수정 가능)
    file_path = "포지션 보고양식 예시.xlsx"

    @st.cache_data
    def load_data():
        sheets = ['9999', '9994', '9992', '9988', '7120(원화)']
        df_list = []
        for s in sheets:
            try:
                temp_df = pd.read_excel(file_path, sheet_name=s)
                df_list.append(temp_df)
            except Exception as e:
                st.warning(f"'{s}' 시트를 찾을 수 없습니다.")
        
        if not df_list: return pd.DataFrame()
        
        data = pd.concat(df_list, ignore_index=True)
        data = data[data['펀드코드'] != 8018].copy() # 8018 제외
        
        # [채권종류 분류]
        def classify_bond(row):
            ctype = str(row['채권종류']).strip()
            name = str(row['종목명']).strip()
            if ctype in ['국채', '통안채', '지방채']: return '국고/통안채'
            if '중앙회' in name or ctype == '특수채': return '특수채'
            if ctype == '금융채':
                if any(x in name for x in ['중소기업은행', '기업은행', '산업금융채권', '산금']): return '특은채'
                if any(x in name for x in ['카드', '캐피탈']): return '여전채'
                if '은행' in name: return '시은채'
                return '기타금융채'
            return '기타'
        
        data['분류_채권종류'] = data.apply(classify_bond, axis=1)
        
        # [잔존만기 분류]
        data['평가일자'] = pd.to_datetime(data['평가일자'])
        data['만기일자'] = pd.to_datetime(data['만기일자'])
        data['잔존년수'] = (data['만기일자'] - data['평가일자']).dt.days / 365.0
        
        def classify_mat(y):
            if y <= 0.25: return '0.25Y'
            elif y <= 0.5: return '0.5Y'
            elif y <= 1.0: return '1Y'
            elif y <= 1.5: return '1.5Y'
            elif y <= 2.0: return '2Y'
            elif y <= 2.5: return '2.5Y'
            elif y <= 3.0: return '3Y'
            elif y <= 5.0: return '5Y'
            elif y <= 10.0: return '10Y'
            elif y <= 20.0: return '20Y'
            else: return '30Y'
            
        data['분류_잔존만기'] = data['잔존년수'].apply(classify_mat)
        
        # [포지션 억원 단위]
        data['포지션(억원)'] = data['수량'] / 100000.0
        
        # [펀드별 카테고리 매핑]
        def map_fund(code):
            if code == 8010: return 'RP운용', '대고객 RP(8010)', '대고객 RP(8010)'
            elif code == 8013: return 'RP운용', 'CMA RP(8013)', 'CMA RP(8013)'
            elif code == 9994: return 'RP운용', '기관RP(9994)', '기관RP(9994)'
            elif code == 7120: return 'RP운용', '외화RP(7120)', '외화RP(7120)'
            elif code in [8001, 8008]: return '상품채권운용', '자격', 'PD(8001)' if code==8001 else '소액(8008)'
            elif code in [8007, 8016, 8011, 8019]:
                name = {8007:'팀운용', 8016:'부서공통', 8011:'Prop1', 8019:'Prop2'}[code]
                return '상품채권운용', '일반 Prop', name
            return '기타', '기타', '기타'

        data[['큰분류', '세부분류', '세세부']] = data.apply(lambda r: pd.Series(map_fund(r['펀드코드'])), axis=1)
        return data

    total_df = load_data()

    if not total_df.empty:
        # 공통 설정
        mat_cols = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']
        
        tab1, tab2 = st.tabs(["🏛️ RP운용", "📈 상품채권운용"])

        with tab1:
            st.subheader("RP운용 포지션 현황")
            rp_order = ['대고객 RP(8010)', 'CMA RP(8013)', '기관RP(9994)', '외화RP(7120)']
            rp_df = total_df[total_df['큰분류'] == 'RP운용'].copy()
            
            if not rp_df.empty:
                rp_df['세부분류'] = pd.Categorical(rp_df['세부분류'], categories=rp_order, ordered=True)
                pivot = pd.pivot_table(rp_df, values='포지션(억원)', 
                                      index=['큰분류', '세부분류', '분류_채권종류'], 
                                      columns='분류_잔존만기', aggfunc='sum', fill_value=0).sort_index(level='세부분류')
                
                # 없는 만기 컬럼 생성 및 정렬
                for c in mat_cols:
                    if c not in pivot.columns: pivot[c] = 0
                pivot = pivot[mat_cols]
                pivot['합계'] = pivot.sum(axis=1)
                
                # HTML 렌더링
                st.markdown(render_merged_html(pivot), unsafe_allow_html=True)
            else:
                st.info("조회된 RP운용 데이터가 없습니다.")

        with tab2:
            st.subheader("상품채권운용 포지션 현황")
            prop_df = total_df[total_df['큰분류'] == '상품채권운용'].copy()
            
            if not prop_df.empty:
                pivot = pd.pivot_table(prop_df, values='포지션(억원)', 
                                      index=['큰분류', '세부분류', '세세부', '분류_채권종류'], 
                                      columns='분류_잔존만기', aggfunc='sum', fill_value=0)
                
                for c in mat_cols:
                    if c not in pivot.columns: pivot[c] = 0
                pivot = pivot[mat_cols]
                pivot['합계'] = pivot.sum(axis=1)
                
                # HTML 렌더링
                st.markdown(render_merged_html(pivot), unsafe_allow_html=True)
            else:
                st.info("조회된 상품채권운용 데이터가 없습니다.")
    else:
        st.error("데이터 로드에 실패했습니다. 파일 경로와 시트명을 확인해 주세요.")
