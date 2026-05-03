# streamlit run .\data-generator.py
import pandas as pd
import streamlit as st
import random

cost_element_ko = [
    ["일반운영경비", "전산비", "전산비", "전산비-시설료"],
    ["일반운영경비", "전산비", "전산비", "전산비-SM-인건비"],
    ["일반운영경비", "전산비", "전산비", "전산비-클라우드"],
    ["일반운영경비", "급여및상여", "상여", "상여-명절상여"],
    ["일반운영경비", "급여및상여", "상여", "상여-특별상여"],
    ["일반운영경비", "감가상각비", "감가상각비", "감가상각비-컴퓨터"],
    ["일반운영경비", "감가상각비", "감가상각비", "감가상각비-차량"],
    ["일반운영경비", "감가상각비", "감가상각비", "감가상각비-비품"],
    ["일반운영경비", "수선비", "수선비-건물", "수선비-건물-청소"],
    ["일반운영경비", "수선비", "수선비-건물", "수선비-건물-도장"],
    ["일반운영경비", "수선비", "수선비-건물", "수선비-건물-수리"],
]

# 비용 코드 할당 로직 (시작 값 5)
@st.cache_data # 데이터가 변경되지 않으면 매번 계산하지 않도록 캐싱
def get_cost_element_mapping(data_list):
    mapping = [{} for _ in range(4)] # Lv1~4 매핑 테이블
    
    for row in data_list:
        for i in range(4): # Lv1부터 Lv4까지 순회
            if row[i] not in mapping[i]:
                if i == 0:
                    # Lv1: 5로 시작하며 순차적 증가 (51, 52, 53...)
                    new_code = f"5{len(mapping[i]) + 1:d}"
                else:
                    # Lv2~4: 01부터 순차적 증가 (01, 02, 03...)
                    new_code = f"{len(mapping[i]) + 1:02d}"
                mapping[i][row[i]] = new_code
    return mapping

@st.cache_data
def get_cost_element_hierarchical_mapping(data_list):
    # mapping[level][parent_key] = {item_name: code}
    mapping = [{} for _ in range(4)]
    
    # Lv1 (고정 5시작)
    lv1_list = sorted(list(set(row[0] for row in data_list)))
    mapping[0] = {name: f"5{i+1:d}" for i, name in enumerate(lv1_list)}
    
    for row in data_list:
        l1, l2, l3, l4 = row
        
        # Lv2 매핑 (부모: Lv1)
        if l1 not in mapping[1]: mapping[1][l1] = {}
        if l2 not in mapping[1][l1]:
            mapping[1][l1][l2] = f"{len(mapping[1][l1]) + 1:02d}"
            
        # Lv3 매핑 (부모: Lv1+Lv2)
        p2 = f"{l1}|{l2}"
        if p2 not in mapping[2]: mapping[2][p2] = {}
        if l3 not in mapping[2][p2]:
            mapping[2][p2][l3] = f"{len(mapping[2][p2]) + 1:02d}"
            
        # Lv4 매핑 (부모: Lv1+Lv2+Lv3)
        p3 = f"{l1}|{l2}|{l3}"
        if p3 not in mapping[3]: mapping[3][p3] = {}
        if l4 not in mapping[3][p3]:
            mapping[3][p3][l4] = f"{len(mapping[3][p3]) + 1:02d}"
            
    return mapping

# cost_element_mapping_table = get_cost_element_mapping(cost_element_ko)
cost_element_mapping_table = get_cost_element_hierarchical_mapping(cost_element_ko)

# 총원장 데이터 생성
def generate_general_ledger_data(fiscal_year, version_id, cost_type, budgeting_type, 
                                 mo_desc, cost_element, functional_area, acct_lv1, 
                                 acct_lv2, acct_lv3, acct_lv4, org_cctr, 
                                 org_cctr_desc, org_job_category, org_bu_desc, mot, 
                                 mot_desc, cost_center, job_category, budget_unit, 
                                 line, vendor, connection_name, org_cost_code, 
                                 org_cost_code_name, item_text, cost_code_type1_5, total_amt, 
                                 sok_12, amt1_12, create_user_id, pic_id, 
                                 record_type, deletion_indicator):
    row = {
        "Fiscal Year": fiscal_year,
        "Version ID": version_id,
        "Cost Type code": cost_type,
        "Budgeting Type": budgeting_type,
        "MO Desc.": mo_desc,
        "Cost Element": cost_element,
        "Functional Area": functional_area,
        "Acct Lv1": acct_lv1,
        "Acct Lv2": acct_lv2,
        "Acct Lv3": acct_lv3,
        "Acct Lv4": acct_lv4,
        "Org. CCTR": org_cctr,
        "Org. CCTR Desc.": org_cctr_desc,
        "Original Job Category": org_job_category,
        "Org. BU Desc.": org_bu_desc,
        "MO(T)": mot,
        "MO(T) Desc.": mot_desc,
        "Cost Center": cost_center,
        "Job Category": job_category,
        "Budget Unit": budget_unit,
        "Line": line,
        "Vendor": vendor,
        "Connection Name": connection_name,
        "Org. Cost Code": org_cost_code,
        "Org. Cost Code_Name": org_cost_code_name,
        "Item Text": item_text,
        "Cost Code Type1~5": cost_code_type1_5,
        "Total Amt": total_amt,
        "속보 12월": sok_12,
        "Amt1~12": amt1_12,
        "Create User ID": create_user_id,
        "PIC ID": pic_id,
        "Record Type": record_type,
        "Deletion Indicator": deletion_indicator
    }
    return row


# 총원장 데이터 생성 UI
st.title("총원장 데이터 생성 (General Ledger Data Generator)")

if 'data_list' not in st.session_state:
    st.session_state['data_list'] = []

# 기본 정보 입력 UI
with st.expander("기본 정보 (Fiscal/Version)"):
    # 회계연도 선택
    fiscal_year = st.selectbox("회계연도", [2022, 2023, 2024, 2025, 2026], index=4)

    # 데이터 구분 선택
    version_type = st.selectbox("데이터 구분", ["실적", "속보", "경영계획"])
    version_id = ""

    # 구분에 따른 세부 규칙 적용
    if version_type == "실적":
        version_id = "000"
        st.info(f"✅ 실적 데이터 버전: **{version_id}**")

    elif version_type == "속보":
        # 속보 규칙: 9~11월(T0A~C), 12월(T01), 1~8월(T02~9)
        month = st.selectbox("대상 월 선택", [f"{i}월" for i in range(1, 13)])
        m_int = int(month.replace("월", ""))
        
        if 9 <= m_int <= 11:
            # 9월->A, 10월->B, 11월->C (ASCII 65가 'A')
            char_code = chr(65 + (m_int - 9))
            version_id = f"T0{char_code}"
        elif m_int == 12:
            version_id = "T01"
        else: # 1~8월
            version_id = f"T0{m_int + 1}"
        
        st.info(f"📅 {month} 속보 버전: **{version_id}**")

    elif version_type == "경영계획":
        # 경영계획 규칙: P01~3 (P03이 최종)
        plan_options = {
            "P01": "P01 (초안)",
            "P02": "P02 (수정안)",
            "P03": "P03 (최종 확정본)"
        }
        selected_plan = st.selectbox(
            "경영계획 차수 선택", 
            options=list(plan_options.keys()),
            format_func=lambda x: plan_options[x],
            index=2  # 기본값을 P03(최종)으로 설정
        )
        version_id = selected_plan
        
        if version_id == "P03":
            st.success(f"🌟 **{plan_options[version_id]}**으로 설정되었습니다.")
        else:
            st.warning(f"⚠️ 현재 선택된 버전은 **{plan_options[version_id]}**입니다.")

# 비용 상세 정보 입력 UI
with st.expander("비용 상세 정보 (Cost Details)"):
    cost_type = st.selectbox("비용의 귀속 (Cost Type code)", ["DIS", "GEN"])
    budgeting_type = st.selectbox("편성주제 (Budgeting Type)", ["EST", "NES"])
    mo_desc = st.text_input("발생사업부명 (MO Desc.)", value="AI센터")
    functional_area = st.selectbox("귀속부서의 성격 (Functional Area)", ["1000", "2000", "3000", "4000", "5000", "6000"])

    with st.container(border=True):
        cols = st.columns(4)
        with cols[0]:
            acct_lv1 = st.selectbox("비용분류 (Acct Lv1)", sorted(list(cost_element_mapping_table[0].keys())))
        with cols[1]:
            acct_lv2 = st.selectbox("(Acct Lv2)", sorted(list(cost_element_mapping_table[1][acct_lv1].keys())))
        with cols[2]:
            acct_lv3 = st.selectbox("(Acct Lv3)", sorted(list(cost_element_mapping_table[2][f"{acct_lv1}|{acct_lv2}"].keys())))
        with cols[3]:
            acct_lv4 = st.selectbox("(Acct Lv4)", sorted(list(cost_element_mapping_table[3][f"{acct_lv1}|{acct_lv2}|{acct_lv3}"].keys())))

        cost_element_code1 = cost_element_mapping_table[0][acct_lv1]
        cost_element_code2 = cost_element_mapping_table[1][acct_lv1][acct_lv2]
        cost_element_code3 = cost_element_mapping_table[2][f"{acct_lv1}|{acct_lv2}"][acct_lv3]
        cost_element_code4 = cost_element_mapping_table[3][f"{acct_lv1}|{acct_lv2}|{acct_lv3}"][acct_lv4]

        cost_element_code = cost_element_code1 + cost_element_code2 + cost_element_code3 + cost_element_code4

        st.success(f"✅ 선택된 비용분류 코드: Lv1({cost_element_code[1:2]}) - Lv2({cost_element_code[2:4]}) - Lv3({cost_element_code[4:6]}) - Lv4({cost_element_code[6:8]})")

        cost_element = st.text_input("비용코드 (Cost Element)", value=cost_element_code, disabled=True)

    org_cctr = st.text_input("발생부서코드 (Org. CCTR)", value="C10F4000")
    org_cctr_desc = st.text_input("발생부서명 (Org. CCTR Desc.)", value="지원팀")
    org_job_category = st.text_input("발생직무코드 (Original Job Category)", value="C21")
    org_bu_desc = st.text_input("B-Hub 기준 귀속조직명 (Org. BU Desc.)", value="지원팀(AI센터)")
    mot = st.text_input("귀속사업부코드 (MO(T))", value="M100B")
    mot_desc = st.text_input("귀속사업부명 (MO(T) Desc.)", value="AI센터")
    cost_center = st.text_input("귀속부서코드 (Cost Center)", value="C10A1000")
    job_category = st.text_input("귀속직무코드 (Job Category)", value="A12")
    budget_unit = st.text_input("B-Hub 기준 귀속조직 (Budget Unit)", value="C100GFA")
    line = st.text_input("(Line)", value="Null")
    vendor = st.text_input("수혜 업체명 (Vendor)", value="Vendor A")
    connection_name = st.text_input("Vendor의 한글명 (Connection Name)", value="벤더 A")
    org_cost_code = st.text_input("발생세부비용코드 (Org. Cost Code)", value="CFA02B000001")
    org_cost_code_name = st.text_input("발생세부비용코드명 (Org. Cost Code_Name)", value="본사배부비_국내통신비")
    item_text = st.text_input("적요 (Item Text)", value="AI센터 국내통신비")
    cost_code_type1_5 = st.text_input("Cost Code에 따른 분류 (Cost Code Type1~5)", value="고정성/시스템 지원/Infra관련/통신비/본사배부")
    total_amt = st.number_input("총 금액 (Total Amt)", min_value=0, value=1000000, step=10000)
    sok_12 = st.number_input("12월속보 편성시 12월금액 (속보 12월)", min_value=0, value=0, step=10000)
    amt1_12 = st.number_input("1-12월 금액 (Amt1~12)", min_value=0, value=1000000, step=10000)
    create_user_id = st.text_input("생성자 (Create User ID)", value="user1234")
    pic_id = st.text_input("Person In Charge (PIC ID)", value="user1234")
    record_type = st.text_input("(Record Type)", value="Null")
    deletion_indicator = st.text_input("(Deletion Indicator)", value="Null")


# 데이터 생성 함수 호출 및 저장
if st.button("데이터 생성"):
    # 데이터 생성 함수 호출
    # st.write(f"최종 전달될 값 -> Fiscal Year: {fiscal_year}, Version Type: {version_type}, Version ID: {version_id} Cost Type code: {cost_type}, Budgeting Type: {budgeting_type}, MO Desc.: {mo_desc}, Cost Element: {cost_element}, Functional Area: {functional_area}, Acct Lv1: {acct_lv1}, Acct Lv2: {acct_lv2}, Acct Lv3: {acct_lv3}, Acct Lv4: {acct_lv4}, Org. CCTR: {org_cctr}, Org. CCTR Desc.: {org_cctr_desc}, Original Job Category: {org_job_category}, Org. BU Desc.: {org_bu_desc}, MO(T): {mot}, MO(T) Desc.: {mot_desc}, Cost Center: {cost_center}, Job Category: {job_category}, Budget Unit: {budget_unit}, Line: {line}, Vendor: {vendor}, Connection Name: {connection_name}, Org. Cost Code: {org_cost_code}, Org. Cost Code_Name: {org_cost_code_name}, Item Text: {item_text}, Cost Code Type1-5: {cost_code_type1_5}, Total Amt: {total_amt}, 속보 12월: {sok_12}, Amt1-12: {amt1_12}, Create User ID: {create_user_id}, PIC ID: {pic_id}, Record Type: {record_type}, Deletion Indicator: {deletion_indicator}")
    new_row = generate_general_ledger_data(
        fiscal_year=fiscal_year,
        version_id=version_id,
        cost_type=cost_type,
        budgeting_type=budgeting_type,
        mo_desc=mo_desc,
        cost_element=cost_element,
        functional_area=functional_area,
        acct_lv1=acct_lv1,
        acct_lv2=acct_lv2,
        acct_lv3=acct_lv3,
        acct_lv4=acct_lv4,
        org_cctr=org_cctr,
        org_cctr_desc=org_cctr_desc,
        org_job_category=org_job_category,
        org_bu_desc=org_bu_desc,
        mot=mot,
        mot_desc=mot_desc,
        cost_center=cost_center,
        job_category=job_category,
        budget_unit=budget_unit,
        line=line,
        vendor=vendor,
        connection_name=connection_name,
        org_cost_code=org_cost_code,
        org_cost_code_name=org_cost_code_name,
        item_text=item_text,
        cost_code_type1_5=cost_code_type1_5,
        total_amt=total_amt,
        sok_12=sok_12,
        amt1_12=amt1_12,
        create_user_id=create_user_id,
        pic_id=pic_id,
        record_type=record_type,
        deletion_indicator=deletion_indicator
    )
    
    st.session_state['data_list'].append(new_row)
    st.success("새로운 데이터가 추가되었습니다.")

# 데이터 출력 및 다운로드
st.divider() # 시각적 구분선
st.subheader("📊 현재 생성된 데이터 내역")

if st.session_state['data_list']:
    # 리스트를 데이터프레임으로 변환
    df = pd.DataFrame(st.session_state['data_list'])
    
    # 결과 미리보기 (st.table은 정적, st.dataframe은 스크롤/정렬 가능)
    st.dataframe(df, use_container_width=True)
    
    # 엑셀(CSV) 다운로드 버튼
    csv_data = df.to_csv(index=False).encode('utf-8-sig')
    
    st.download_button(
        label="📥 생성된 데이터를 엑셀(CSV)로 다운로드",
        data=csv_data,
        file_name="gl_test_data.csv",
        mime="text/csv",
        help="클릭하면 지금까지 생성된 모든 데이터를 파일로 저장합니다."
    )
else:
    st.info("아직 생성된 데이터가 없습니다.")
