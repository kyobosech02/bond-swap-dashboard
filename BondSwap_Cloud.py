import streamlit as st
import pandas as pd
import numpy as np

# 페이지 설정
st.set_page_config(layout="wide", page_title="포지션 및 델타 현황")

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

if check_password():
    st.title("원화채권 포지션 및 델타 현황")

    file_path = "포지션 보고양식 예시.xlsx"

    # 고정 순서 정의
    BOND_ORDER = ['국고/통안채', '특은채', '시은채', '특수채', '여전채', '기타금융채']
    MATURITY_ORDER = ['0.25Y', '0.5Y', '1Y', '1.5Y', '2Y', '2.5Y', '3Y', '5Y', '10Y', '20Y', '30Y']

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
        
        # 1. 채권종류 분류
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
        
        # 2. 잔존만기 분류
        data['평가일자'] = pd.to_datetime(data['평가일자'])
        data['만기일자'] = pd.to_datetime(data['만기일자'])
        data['잔존년수'] = (data['만기일자'] - data['평가일자']).dt.days / 365.0
        
        def classify_maturity(years):
            bins = [0, 0.25, 0.5, 1, 1.5, 2, 2.5, 3, 5, 10, 20, 1000]
            labels = MATURITY_ORDER
            for i, b in enumerate(bins[1:]):
                if years <= b: return labels[i]
            return '30Y'
        data['분류_잔존만기'] = data['잔존년수'].apply(classify_maturity)
        
        # 3. 데이터 계산 (포지션 억원, 델타 만원)
        data['포지션'] = data['수량'] / 100000.0
        data['델타'] = (data['평가금액'] * data['수정듀레이션']) / 100000000.0
        
        # 4. 펀드별 카테고리 매핑
        def classify_fund(code):
            if code == 8010: return 'RP운용', '대고객 RP(8010)', '대고객 RP(8010)'
            elif code == 8013: return 'RP운용', 'CMA RP(8013)', 'CMA RP(8013)'
            elif code == 9994: return 'RP운용', '기관RP(9994)', '기관RP(9994)'
            elif code == 7120: return 'RP운용', '외화RP(7120)', '외화RP(7120)'
            elif code in [8001, 8008]: return '상품채권운용', '자격', 'PD(8001)' if code==8001 else '소액(8008)'
            elif code in [8007, 8016, 8011, 8019]:
                names = {8007:'팀운용(8007)', 8016:'부서공통(8016)', 8011:'Prop1(8011)', 8019:'Prop2(8019)'}
                return '상품채권운용', '일반 Prop', names[code]
            return '기타', '기타', '기타'
        data[['큰분류', '세부분류', '세세부']] = data.apply(lambda row: pd.Series(classify_fund(row['펀드코드'])), axis=1)
        return data

    def process_pivot_with_fixed_order(df, index_cols, target_col):
        """값이 없어도 모든 채권종류를 고정 순서로 생성하고 소계를 추가하는 함수"""
        # 피벗 생성
        pivot = pd.pivot_table(df, values=target_col, index=index_cols + ['분류_채권종류'], 
                               columns='분류_잔존만기', aggfunc='sum', fill_value=0)
        
        # 모든 만기 컬럼 확보
        for m in MATURITY_ORDER:
            if m not in pivot.columns: pivot[m] = 0
        pivot = pivot[MATURITY_ORDER]

        # 모든 조합 생성 (Reindexing)
        # 상위 인덱스 레벨들의 유니크 값 추출
        unique_levels = [pivot.index.get_level_values(i).unique() for i in range(len(index_cols))]
        # 채권종류는 BOND_ORDER로 고정
        unique_levels.append(BOND_ORDER)
        
        full_index = pd.MultiIndex.from_product(unique_levels, names=pivot.index.names)
        pivot = pivot.reindex(full_index, fill_value=0)

        # 소계 계산 및 삽입
        final_rows = []
        # 상위 그룹별로 순회하며 소계 추가
        for name, group in pivot.groupby(level=list(range(len(index_cols)))):
            final_rows.append(group)
            # 소계 행 생성
            subtotal_row = group.sum().to_frame().T
            # 소계 인덱스 설정 (마지막 레벨을 '소계'로)
            sub_idx = list(name) if isinstance(name, tuple) else [name]
            sub_idx.append('소계')
            subtotal_row.index = pd.MultiIndex.from_tuples([tuple(sub_idx)], names=pivot.index.names)
            final_rows.append(subtotal_row)
            
        result = pd.concat(final_rows)
        result['합계'] = result.sum(axis=1)
        return result

    def style_dataframe(df):
        return df.style.format("{:,.0f}") \
            .set_properties(**{'border': '1px solid black', 'text-align': 'center'}) \
            .apply(lambda x: ['background-color: #f2f2f2; font-weight: bold' if '소계' in str(x.name) else '' for _ in x], axis=1) \
            .set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#e0e0e0'), ('border', '2px solid black'), ('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('border', '1px solid #444')]}
            ])

    df = load_and_preprocess_data()

    if not df.empty:
        st.write("### 🔍 보기 설정")
        view_type = st.radio("표시 데이터 선택", ["포지션 (단위: 억원)", "델타 (단위: 만원)"], horizontal=True)
        target_col = "포지션" if "포지션" in view_type else "델타"
        
        rp_order = ['대고객 RP(8010)', 'CMA RP(8013)', '기관RP(9994)', '외화RP(7120)']
        
        tab1, tab2 = st.tabs([f"🏛️ RP운용 ({target_col})", f"📈 상품채권운용 ({target_col})"])

        with tab1:
            rp_df = df[df['큰분류'] == 'RP운용'].copy()
            if not rp_df.empty:
                # RP 세부분류 순서 강제 지정을 위해 Categorical 사용
                rp_df['세부분류'] = pd.Categorical(rp_df['세부분류'], categories=rp_order, ordered=True)
                
                # 피벗 및 소계 처리 (큰분류, 세부분류 기준)
                pivot = process_pivot_with_fixed_order(rp_df, ['큰분류', '세부분류'], target_col)
                
                # Categorical 순서에 따라 정렬
                pivot = pivot.sort_index(level='세부분류', sort_remaining=False)
                
                h = (len(pivot) + 1) * 38 + 50
                st.dataframe(style_dataframe(pivot), use_container_width=True, height=int(h))
            else:
                st.info("데이터가 없습니다.")

        with tab2:
            prop_df = df[df['큰분류'] == '상품채권운용'].copy()
            if not prop_df.empty:
                # 피벗 및 소계 처리 (큰분류, 세부분류, 세세부 기준)
                pivot = process_pivot_with_fixed_order(prop_df, ['큰분류', '세부분류', '세세부'], target_col)
                
                h = (len(pivot) + 1) * 38 + 50
                st.dataframe(style_dataframe(pivot), use_container_width=True, height=int(h))
            else:
                st.info("데이터가 없습니다.")
    else:
        st.error("데이터 로드 실패")
