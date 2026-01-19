"""
🏢 HWPX 파일 동별/층별 분석 웹 앱
- HWPX 파일 업로드 → 동별 분리 → 층별 분리 → 엑셀 다운로드
"""
import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
import re
import io

# ===== HWPX 파싱 함수들 =====
NS = {'hwp': 'http://www.hancom.co.kr/hwpml/2011/paragraph'}

def find_elements(parent, tag):
    """네임스페이스 있든 없든 요소 찾기"""
    elements = parent.findall(f'.//{tag}')
    if not elements:
        elements = parent.findall(f'.//hwp:{tag}', NS)
    return elements

def get_text(element):
    """요소에서 모든 텍스트 추출"""
    texts = [t.text for t in find_elements(element, 't') if t.text]
    return ' '.join(texts)

def extract_table_data(table):
    """표를 2차원 리스트로 변환"""
    rows_data = []
    for row in find_elements(table, 'tr'):
        row_data = [get_text(cell).strip() for cell in find_elements(row, 'tc')]
        if row_data:
            rows_data.append(row_data)
    return rows_data

def extract_tables_from_hwpx(hwpx_path):
    """HWPX 파일에서 모든 표 추출"""
    all_tables = []
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        for section_file in [f for f in z.namelist() if f.startswith('Contents/section')]:
            with z.open(section_file) as f:
                root = ET.parse(f).getroot()
                for table in find_elements(root, 'tbl'):
                    table_data = extract_table_data(table)
                    if table_data:
                        all_tables.append(table_data)
    return all_tables

def group_by_dong(all_tables):
    """표 데이터를 동별로 그룹화"""
    dong_data = {}
    current_dong = None
    for table in all_tables:
        for row in table:
            match = re.search(r'(\d+)동', ' '.join(row))
            if match:
                current_dong = f"{match.group(1)}동"
                if current_dong not in dong_data:
                    dong_data[current_dong] = []
            if current_dong:
                dong_data[current_dong].append(row)
    return dong_data

def filter_by_floor_range(rows, start_floor, end_floor):
    """특정 층 범위에 해당하는 행만 필터링"""
    filtered_rows = []
    for row in rows:
        row_text = ' '.join([str(cell) for cell in row if cell])
        # 숫자층 검색
        for floor_num in range(start_floor, end_floor + 1):
            if f"{floor_num}층" in row_text or f"[ {floor_num}층" in row_text:
                filtered_rows.append(row)
                break
        # 옥탑/지붕층 (15층 이상)
        if end_floor >= 15 and any(kw in row_text for kw in ["옥탑", "지붕"]):
            if row not in filtered_rows:
                filtered_rows.append(row)
    return filtered_rows

# ===== Streamlit 앱 시작 =====

st.set_page_config(page_title="HWPX 분석기", layout="wide")

st.title("🏢 HWPX 파일 동별/층별 분석기")

# 세션 스테이트 초기화
if 'dong_data' not in st.session_state:
    st.session_state.dong_data = {}
if 'floor_ranges' not in st.session_state:
    st.session_state.floor_ranges = {}

# 파일 업로드
uploaded_file = st.file_uploader("HWPX 파일 업로드", type=['hwpx'])

if uploaded_file:
    with st.spinner("파일 분석 중..."):
        # 표 추출
        tables_data = extract_tables_from_hwpx(uploaded_file)
        st.success(f"✅ {len(tables_data)}개의 표를 발견했습니다")
        
        # 동별로 파싱
        dong_data = group_by_dong(tables_data)
        st.session_state.dong_data = dong_data
        
        st.info(f"📊 총 {len(dong_data)}개의 동을 발견했습니다: {', '.join(sorted(dong_data.keys()))}")
    
    # 동 선택
    selected_dong = st.selectbox("동 선택", sorted(dong_data.keys(), key=lambda x: int(re.search(r'\d+', x).group())))
    
    if selected_dong:
        dong_rows = dong_data[selected_dong]
        
        st.subheader(f"{selected_dong} - 층별 범위 설정")
        
        # 데이터 미리보기 (디버깅용)
        with st.expander("🔍 원본 데이터 미리보기"):
            df_preview = pd.DataFrame(dong_rows)
            st.dataframe(df_preview.head(100), use_container_width=True)
            
            # 첫 번째 셀 값들만 확인
            st.markdown("**첫 번째 열 값들:**")
            first_col_values = [str(row[0]) if row and len(row) > 0 else "" for row in dong_rows[:50]]
            for i, val in enumerate(first_col_values):
                if val.strip():
                    st.text(f"{i}: {val}")
        
        # 층 추가
        with st.form(f"add_floor_{selected_dong}"):
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                floor_name = st.text_input("층 이름", placeholder="예: 저층부")
            with col2:
                floor_type = st.selectbox("구분", ["지상", "지하"])
            with col3:
                start_floor = st.number_input("시작 층", min_value=1, max_value=100, value=1)
            with col4:
                end_floor = st.number_input("끝 층", min_value=1, max_value=100, value=8)
            
            submitted = st.form_submit_button("층 추가")
            
            if submitted and floor_name:
                if selected_dong not in st.session_state.floor_ranges:
                    st.session_state.floor_ranges[selected_dong] = {}
                
                st.session_state.floor_ranges[selected_dong][floor_name] = {
                    'floor_type': floor_type,
                    'start_floor': int(start_floor),
                    'end_floor': int(end_floor)
                }
                floor_display = f"{floor_type} {start_floor}층 ~ {end_floor}층"
                st.success(f"✅ {floor_name} 추가됨 ({floor_display})")
                st.rerun()
        
        # 설정된 층 범위 표시
        if selected_dong in st.session_state.floor_ranges:
            st.markdown("### 설정된 층 범위")
            
            for floor_name, range_info in st.session_state.floor_ranges[selected_dong].items():
                col1, col2 = st.columns([4, 1])
                with col1:
                    floor_type = range_info.get('floor_type', '지상')
                    st.write(f"**{floor_name}**: {floor_type} {range_info['start_floor']}층 ~ {range_info['end_floor']}층")
                with col2:
                    if st.button("삭제", key=f"del_{selected_dong}_{floor_name}"):
                        del st.session_state.floor_ranges[selected_dong][floor_name]
                        st.rerun()

# 엑셀 다운로드
st.markdown("---")
st.subheader("💾 엑셀 파일 생성")

def make_excel_bytes():
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for dong_name in sorted(
            st.session_state.dong_data.keys(),
            key=lambda x: int(re.search(r'\d+', x).group())
        ):
            rows = st.session_state.dong_data[dong_name]
            df = pd.DataFrame(rows)

            if dong_name in st.session_state.floor_ranges:
                combined_data = []

                for floor_name, range_info in st.session_state.floor_ranges[dong_name].items():
                    floor_type = range_info.get('floor_type', '지상')
                    start_floor = range_info['start_floor']
                    end_floor = range_info['end_floor']

                    # 제목 행
                    title_row = [f"[ {floor_name} ]"] + [''] * (df.shape[1] - 1)
                    combined_data.append(title_row)

                    all_rows = df.values.tolist()

                    for floor_num in range(start_floor, end_floor + 1):
                        floor_rows = []
                        capturing = False
                        skip_section = False

                        for row in all_rows:
                            row_text = ' '.join([str(cell) if cell is not None else '' for cell in row])

                            # 정면도/배면도 스킵
                            if '정면도' in row_text or '배면도' in row_text:
                                skip_section = True
                                continue

                            if re.search(r'\d+동\s*\d+층|\d+동\s*(지하|B)\s*\d+층', row_text):
                                skip_section = False

                            if skip_section:
                                continue

                            # 제외 키워드 필터
                            exclude_keywords = [
                                '부록','외관조사망도','참조','번 호',
                                '부   위','부 재','폭','mm','길이','개소','EA'
                            ]
                            if any(k in row_text for k in exclude_keywords):
                                continue

                            if floor_type == "지하":
                                if re.search(rf'\d+동\s*(지하|B)\s*{floor_num}층', row_text):
                                    capturing = True
                                    floor_rows.append(row)
                                    continue

                                if capturing and re.search(r'\d+동\s*\d+층|\d+동\s*(지하|B)\s*\d+층', row_text):
                                    break

                            else:
                                match = re.search(rf'\d+동\s*(\d+)층', row_text)
                                if match and "지하" not in row_text and "B" not in row_text:
                                    if int(match.group(1)) == floor_num:
                                        capturing = True
                                        floor_rows.append(row)
                                        continue

                                if capturing and re.search(r'\d+동\s*\d+층|\d+동\s*(지하|B)\s*\d+층', row_text):
                                    break

                            if capturing:
                                floor_rows.append(row)

                        if floor_rows:
                            combined_data.extend(floor_rows)

                combined_df = pd.DataFrame(combined_data)
                combined_df.to_excel(writer, sheet_name=dong_name[:31], index=False, header=False)

    output.seek(0)
    return output.getvalue()

# ===== 다운로드 버튼 스타일 (여기에 넣기) =====
st.markdown(
    """
    <style>
    /* Streamlit 다운로드 버튼을 .btnBlueGreen.btnPush 스타일로 매핑 */
    div.stDownloadButton > button {
        color: white;
        background: #00AE68;
        box-shadow: 0px 5px 0px 0px #007144;
        transition: all 0.1s ease-in-out;
    }

    /* .btnPush:hover 에 해당 */
    div.stDownloadButton > button:hover {
        color: white;
        background: #00AE68;    
        margin-top: 5px;
        margin-bottom: 5px;
        box-shadow: 0px 0px 0px 0px #007144;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ===== 기존 다운로드 버튼 (그대로 두기) =====
if st.session_state.dong_data and st.session_state.floor_ranges:
    st.download_button(
        label="📥 엑셀 다운로드",
        data=make_excel_bytes(),
        file_name="층별_분석.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
else:
    st.warning("⚠️ 층별 범위를 먼저 설정해주세요")


# 사이드바
with st.sidebar:
    st.header("📋 설정된 모든 층 범위")
    
    if st.session_state.floor_ranges:
        for dong_name, floors in st.session_state.floor_ranges.items():
            st.subheader(dong_name)
            for floor_name, range_info in floors.items():
                st.write(f"• {floor_name}: {range_info['start_floor']}~{range_info['end_floor']}층")
    else:
        st.info("아직 설정된 층 범위가 없습니다")
    
    if st.button("🔄 모두 초기화", use_container_width=True):
        st.session_state.floor_ranges = {}
        st.session_state.dong_data = {}
        st.rerun()
