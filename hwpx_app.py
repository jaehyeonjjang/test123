import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import re
import io

st.set_page_config(page_title="HWPX 분석기", layout="wide")

st.title("🏢 HWPX 파일 동별/층별 분석기")

# 세션 스테이트 초기화
if 'dong_data' not in st.session_state:
    st.session_state.dong_data = {}
if 'floor_ranges' not in st.session_state:
    st.session_state.floor_ranges = {}

def extract_text_from_element(element):
    """XML 요소에서 모든 텍스트 추출"""
    texts = []
    for t in element.findall('.//t'):
        if t.text:
            texts.append(t.text)
    for t in element.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}t'):
        if t.text:
            texts.append(t.text)
    return ' '.join(texts)

def extract_table_data(table_element):
    """표 요소에서 데이터 추출"""
    rows_data = []
    rows = table_element.findall('.//tr')
    if not rows:
        rows = table_element.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tr')
    
    for row in rows:
        row_data = []
        cells = row.findall('.//tc')
        if not cells:
            cells = row.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tc')
        
        for cell in cells:
            text = extract_text_from_element(cell)
            row_data.append(text.strip())
        
        if row_data:
            rows_data.append(row_data)
    
    return rows_data

def extract_tables_from_hwpx(hwpx_file):
    """HWPX 파일에서 표 데이터 추출"""
    tables_data = []
    
    with zipfile.ZipFile(hwpx_file, 'r') as z:
        section_files = [f for f in z.namelist() if f.startswith('Contents/section')]
        
        for section_file in section_files:
            with z.open(section_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
                tables = root.findall('.//tbl')
                if not tables:
                    tables = root.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tbl')
                
                for idx, table in enumerate(tables):
                    table_data = extract_table_data(table)
                    if table_data:
                        tables_data.append(table_data)
    
    return tables_data

def parse_data_by_dong(tables_data):
    """표 데이터를 동별로 파싱"""
    dong_pattern = r'(\d+)동'
    dong_sections = {}
    current_dong = None
    
    for table in tables_data:
        for row in table:
            for cell in row:
                dong_match = re.search(dong_pattern, cell)
                if dong_match:
                    current_dong = f"{dong_match.group(1)}동"
                    if current_dong not in dong_sections:
                        dong_sections[current_dong] = []
            
            if current_dong:
                dong_sections[current_dong].append(row)
    
    return dong_sections

# 파일 업로드
uploaded_file = st.file_uploader("HWPX 파일 업로드", type=['hwpx'])

if uploaded_file:
    with st.spinner("파일 분석 중..."):
        # 표 추출
        tables_data = extract_tables_from_hwpx(uploaded_file)
        st.success(f"✅ {len(tables_data)}개의 표를 발견했습니다")
        
        # 동별로 파싱
        dong_data = parse_data_by_dong(tables_data)
        st.session_state.dong_data = dong_data
        
        st.info(f"📊 총 {len(dong_data)}개의 동을 발견했습니다: {', '.join(sorted(dong_data.keys()))}")
    
    # 동 선택
    selected_dong = st.selectbox("동 선택", sorted(dong_data.keys(), key=lambda x: int(re.search(r'\d+', x).group())))
    
    if selected_dong:
        dong_rows = dong_data[selected_dong]
        
        # DataFrame 생성 (미리보기용)
        df = pd.DataFrame(dong_rows)
        
        st.subheader(f"{selected_dong} 데이터 ({len(dong_rows)}개 행)")
        
        # 데이터 미리보기
        st.dataframe(df.head(50), use_container_width=True)
        
        st.markdown("---")
        st.subheader("📐 층별 범위 설정")
        
        # 층 추가
        with st.form(f"add_floor_{selected_dong}"):
            col1, col2, col3 = st.columns(3)
            with col1:
                floor_name = st.text_input("층 이름", placeholder="예: 저층부(1~8층)")
            with col2:
                start_floor = st.number_input("시작 층", min_value=1, max_value=100, value=1)
            with col3:
                end_floor = st.number_input("끝 층", min_value=1, max_value=100, value=8)
            
            submitted = st.form_submit_button("층 추가")
            
            if submitted and floor_name:
                if selected_dong not in st.session_state.floor_ranges:
                    st.session_state.floor_ranges[selected_dong] = {}
                
                st.session_state.floor_ranges[selected_dong][floor_name] = {
                    'start_floor': int(start_floor),
                    'end_floor': int(end_floor)
                }
                st.success(f"✅ {floor_name} 추가됨 ({start_floor}층 ~ {end_floor}층)")
                st.rerun()
        
        # 설정된 층 범위 표시
        if selected_dong in st.session_state.floor_ranges:
            st.markdown("### 설정된 층 범위")
            
            for floor_name, range_info in st.session_state.floor_ranges[selected_dong].items():
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.write(f"**{floor_name}**: {range_info['start_floor']}층 ~ {range_info['end_floor']}층")
                with col2:
                    if st.button("삭제", key=f"del_{selected_dong}_{floor_name}"):
                        del st.session_state.floor_ranges[selected_dong][floor_name]
                        st.rerun()

# 엑셀 다운로드
st.markdown("---")
st.subheader("💾 엑셀 파일 생성")

col1, col2 = st.columns(2)

with col1:
    if st.button("동별로만 분리", type="secondary", use_container_width=True):
        if st.session_state.dong_data:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for dong_name in sorted(st.session_state.dong_data.keys(), key=lambda x: int(re.search(r'\d+', x).group())):
                    rows = st.session_state.dong_data[dong_name]
                    df = pd.DataFrame(rows)
                    df.to_excel(writer, sheet_name=dong_name[:31], index=False, header=False)
            
            st.download_button(
                label="📥 동별 엑셀 다운로드",
                data=output.getvalue(),
                file_name="동별_분석.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

with col2:
    if st.button("층별로도 분리", type="primary", use_container_width=True):
        if st.session_state.dong_data and st.session_state.floor_ranges:
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for dong_name in sorted(st.session_state.dong_data.keys(), key=lambda x: int(re.search(r'\d+', x).group())):
                    rows = st.session_state.dong_data[dong_name]
                    df = pd.DataFrame(rows)
                    
                    # 층 설정이 있으면 층별로 분리
                    if dong_name in st.session_state.floor_ranges:
                        combined_data = []
                        
                        for floor_name, range_info in st.session_state.floor_ranges[dong_name].items():
                            # 층 제목
                            title_row = [f"[ {floor_name} ]"] + [''] * (df.shape[1] - 1)
                            combined_data.append(title_row)
                            
                            # 해당 층 범위에 속하는 행 찾기
                            start_floor = range_info['start_floor']
                            end_floor = range_info['end_floor']
                            
                            floor_rows = []
                            for idx, row in df.iterrows():
                                row_text = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                                
                                # 숫자층 패턴 (1층, 2층, ...)
                                for floor_num in range(start_floor, end_floor + 1):
                                    if f"{floor_num}층" in row_text or f"[ {floor_num}층" in row_text:
                                        floor_rows.append(row.tolist())
                                        break
                                
                                # 특수층 패턴 (옥탑, 지붕 등) - 층 범위가 높은 경우
                                if end_floor >= 15:  # 15층 이상이면 옥탑/지붕층도 포함
                                    if any(keyword in row_text for keyword in ["옥탑", "지붕"]):
                                        if row.tolist() not in floor_rows:
                                            floor_rows.append(row.tolist())
                            
                            # 데이터 추가
                            combined_data.extend(floor_rows)
                            
                            # 빈 행
                            combined_data.append([''] * df.shape[1])
                        
                        combined_df = pd.DataFrame(combined_data)
                        combined_df.to_excel(writer, sheet_name=dong_name[:31], index=False, header=False)
                    else:
                        # 층 설정 없으면 원본 그대로
                        df.to_excel(writer, sheet_name=dong_name[:31], index=False, header=False)
            
            st.download_button(
                label="📥 층별 엑셀 다운로드",
                data=output.getvalue(),
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
