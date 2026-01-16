import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import re
import io

def extract_tables_from_hwpx(hwpx_file):
    """HWPX 파일에서 표 데이터 추출"""
    tables_data = []
    
    try:
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
                            tables_data.append({
                                'section': section_file,
                                'table_index': idx,
                                'data': table_data
                            })
    except Exception as e:
        st.error(f"오류 발생: {e}")
    
    return tables_data

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

def group_data_by_dong(tables_data):
    """표 데이터를 동별로 그룹화"""
    dong_data = {}
    current_dong = None
    
    for table_info in tables_data:
        table = table_info['data']
        
        for row in table:
            for cell in row:
                dong_match = re.search(r'(\d+)동', cell)
                if dong_match:
                    current_dong = f"{dong_match.group(1)}동"
                    if current_dong not in dong_data:
                        dong_data[current_dong] = []
            
            if current_dong:
                dong_data[current_dong].append(row)
    
    return dong_data

def main():
    st.set_page_config(page_title="HWPX 층별 분석기", layout="wide")
    
    st.title("🏢 HWPX 파일 층별 데이터 분석기")
    st.markdown("HWPX 파일을 업로드하고 동별, 층별로 데이터를 분리하여 엑셀로 저장하세요.")
    
    # 세션 상태 초기화
    if 'dong_data' not in st.session_state:
        st.session_state.dong_data = None
    if 'floor_ranges' not in st.session_state:
        st.session_state.floor_ranges = {}
    
    # 파일 업로드
    st.sidebar.header("📁 파일 선택")
    
    # 로컬 파일 경로 또는 업로드 선택
    file_source = st.sidebar.radio("파일 소스:", ["로컬 경로", "파일 업로드"])
    
    hwpx_file = None
    
    if file_source == "로컬 경로":
        hwpx_path = st.sidebar.text_input(
            "HWPX 파일 경로:",
            value="/Users/seongjaehyeon/project/hwpx/00. 본보고서 - 동탄더레이크팰리스.hwpx"
        )
        if hwpx_path and Path(hwpx_path).exists():
            hwpx_file = hwpx_path
        elif hwpx_path:
            st.sidebar.error("파일을 찾을 수 없습니다.")
    else:
        uploaded_file = st.sidebar.file_uploader("HWPX 파일 선택", type=['hwpx'])
        if uploaded_file:
            hwpx_file = uploaded_file
    
    # 파일 처리
    if hwpx_file and st.sidebar.button("📊 파일 분석 시작", type="primary"):
        with st.spinner("HWPX 파일 분석 중..."):
            tables_data = extract_tables_from_hwpx(hwpx_file)
            st.session_state.dong_data = group_data_by_dong(tables_data)
            st.success(f"✅ 분석 완료! {len(st.session_state.dong_data)}개 동 발견")
    
    # 동별 데이터가 있을 때
    if st.session_state.dong_data:
        dong_data = st.session_state.dong_data
        
        st.header("🏗️ 동별 데이터")
        
        # 동 목록
        dong_list = sorted(dong_data.keys(), key=lambda x: int(re.search(r'\d+', x).group()))
        
        # 탭으로 동별 구분
        tabs = st.tabs(dong_list)
        
        for idx, dong_name in enumerate(dong_list):
            with tabs[idx]:
                rows = dong_data[dong_name]
                
                st.subheader(f"{dong_name} (총 {len(rows)}개 행)")
                
                # 데이터 미리보기
                df_preview = pd.DataFrame(rows)
                
                # 행 번호 표시를 위해 인덱스 조정
                df_preview.index = range(1, len(df_preview) + 1)
                
                st.dataframe(df_preview, height=300, use_container_width=True)
                
                # 층별 범위 설정
                st.markdown("### 📐 층별 범위 설정")
                
                if dong_name not in st.session_state.floor_ranges:
                    st.session_state.floor_ranges[dong_name] = []
                
                # 기존 층 범위 표시
                floor_configs = st.session_state.floor_ranges[dong_name]
                
                # 새 층 추가
                col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
                
                with col1:
                    floor_name = st.text_input(
                        "층 이름",
                        key=f"floor_name_{dong_name}",
                        placeholder="예: 옥탑 2층, 15층"
                    )
                with col2:
                    start_row = st.number_input(
                        "시작 행",
                        min_value=1,
                        max_value=len(rows),
                        value=1,
                        key=f"start_{dong_name}"
                    )
                with col3:
                    end_row = st.number_input(
                        "끝 행",
                        min_value=1,
                        max_value=len(rows),
                        value=min(10, len(rows)),
                        key=f"end_{dong_name}"
                    )
                with col4:
                    if st.button("➕ 추가", key=f"add_{dong_name}"):
                        if floor_name:
                            st.session_state.floor_ranges[dong_name].append({
                                'name': floor_name,
                                'start': start_row,
                                'end': end_row
                            })
                            st.rerun()
                        else:
                            st.warning("층 이름을 입력하세요.")
                
                # 설정된 층 목록
                if floor_configs:
                    st.markdown("#### 📋 설정된 층 범위")
                    for i, config in enumerate(floor_configs):
                        col1, col2, col3 = st.columns([3, 2, 1])
                        with col1:
                            st.text(f"{config['name']}")
                        with col2:
                            st.text(f"{config['start']}행 ~ {config['end']}행")
                        with col3:
                            if st.button("🗑️", key=f"del_{dong_name}_{i}"):
                                st.session_state.floor_ranges[dong_name].pop(i)
                                st.rerun()
        
        # 엑셀 저장
        st.header("💾 엑셀 저장")
        
        col1, col2 = st.columns(2)
        
        with col1:
            output_filename = st.text_input(
                "출력 파일명",
                value="층별분석결과.xlsx"
            )
        
        with col2:
            st.write("")  # 간격
            st.write("")  # 간격
            if st.button("💾 엑셀로 저장", type="primary", use_container_width=True):
                # 엑셀 생성
                output = io.BytesIO()
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for dong_name in dong_list:
                        rows = dong_data[dong_name]
                        df = pd.DataFrame(rows)
                        
                        # 층별 범위가 설정된 경우
                        if dong_name in st.session_state.floor_ranges and st.session_state.floor_ranges[dong_name]:
                            for floor_config in st.session_state.floor_ranges[dong_name]:
                                start_idx = floor_config['start'] - 1
                                end_idx = floor_config['end']
                                
                                floor_df = df.iloc[start_idx:end_idx].copy()
                                sheet_name = f"{dong_name}_{floor_config['name']}"[:31]
                                floor_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        else:
                            # 범위 미설정 시 전체 저장
                            df.to_excel(writer, sheet_name=dong_name[:31], index=False, header=False)
                
                output.seek(0)
                
                st.download_button(
                    label="⬇️ 파일 다운로드",
                    data=output,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success("✅ 엑셀 파일이 준비되었습니다!")

if __name__ == "__main__":
    main()
