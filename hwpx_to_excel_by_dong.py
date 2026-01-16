import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
from pathlib import Path
import re

def extract_tables_from_hwpx(hwpx_path):
    """HWPX 파일에서 표 데이터 추출"""
    tables_data = []
    
    try:
        with zipfile.ZipFile(hwpx_path, 'r') as z:
            # HWPX 파일 구조 확인
            print(f"\nHWPX 파일 내부 구조:")
            for name in z.namelist()[:20]:  # 처음 20개만 출력
                print(f"  - {name}")
            
            # Contents 폴더 내의 섹션 파일들 찾기
            section_files = [f for f in z.namelist() if f.startswith('Contents/section')]
            
            print(f"\n발견된 섹션 파일: {len(section_files)}개")
            
            for section_file in section_files:
                print(f"\n처리 중: {section_file}")
                
                # XML 파싱
                with z.open(section_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    
                    # 네임스페이스 처리
                    namespaces = {'hwp': 'http://www.hancom.co.kr/hwpml/2011/paragraph'}
                    
                    # 표 찾기 (tbl 태그)
                    tables = root.findall('.//hwp:tbl', namespaces)
                    
                    if not tables:
                        # 네임스페이스 없이 시도
                        tables = root.findall('.//tbl')
                    
                    print(f"  발견된 표: {len(tables)}개")
                    
                    for idx, table in enumerate(tables):
                        table_data = extract_table_data(table)
                        if table_data:
                            tables_data.append({
                                'section': section_file,
                                'table_index': idx,
                                'data': table_data
                            })
    
    except Exception as e:
        print(f"오류 발생: {e}")
    
    return tables_data

def extract_table_data(table_element):
    """표 요소에서 데이터 추출"""
    rows_data = []
    
    # 모든 행(tr) 찾기
    rows = table_element.findall('.//tr')
    if not rows:
        rows = table_element.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tr')
    
    for row in rows:
        row_data = []
        
        # 각 행의 셀(tc) 찾기
        cells = row.findall('.//tc')
        if not cells:
            cells = row.findall('.//{http://www.hancom.co.kr/hwpml/2011/paragraph}tc')
        
        for cell in cells:
            # 셀의 텍스트 추출
            text = extract_text_from_element(cell)
            row_data.append(text.strip())
        
        if row_data:
            rows_data.append(row_data)
    
    return rows_data

def extract_text_from_element(element):
    """XML 요소에서 모든 텍스트 추출"""
    texts = []
    
    # t 태그에서 텍스트 찾기
    for t in element.findall('.//t'):
        if t.text:
            texts.append(t.text)
    
    # 네임스페이스 포함해서 다시 시도
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
            # 동 패턴 찾기
            for cell in row:
                dong_match = re.search(r'(\d+)동', cell)
                if dong_match:
                    current_dong = f"{dong_match.group(1)}동"
                    if current_dong not in dong_data:
                        dong_data[current_dong] = []
                    print(f"동 발견: {current_dong}")
            
            # 현재 동이 설정되어 있으면 데이터 추가
            if current_dong:
                dong_data[current_dong].append(row)
    
    return dong_data

def main():
    # HWPX 파일 찾기
    project_path = Path('/Users/seongjaehyeon/project')
    hwpx_files = list(project_path.glob('*.hwpx'))
    hwpx_files.extend(list(project_path.glob('hwpx/*.hwpx')))
    
    if not hwpx_files:
        print("HWPX 파일을 찾을 수 없습니다.")
        print("\n💡 HWP 파일을 한글 프로그램에서 HWPX 형식으로 저장해주세요:")
        print("   1. HWP 파일 열기")
        print("   2. 파일 > 다른 이름으로 저장")
        print("   3. 파일 형식을 'HWPX(한글 2007 문서)'로 선택")
        return
    
    print(f"발견된 HWPX 파일: {len(hwpx_files)}개")
    for f in hwpx_files:
        print(f"  - {f.name}")
    
    # 동탄더레이크팰리스 파일 찾기
    target_file = None
    for f in hwpx_files:
        if '동탄더레이크팰리스' in f.name or '동탄' in f.name:
            target_file = f
            break
    
    if not target_file:
        target_file = hwpx_files[0]
    
    print(f"\n처리할 파일: {target_file.name}")
    
    # 표 데이터 추출
    tables_data = extract_tables_from_hwpx(target_file)
    
    print(f"\n총 추출된 표: {len(tables_data)}개")
    
    # 표 내용 미리보기
    if tables_data:
        print("\n=== 첫 번째 표 미리보기 ===")
        first_table = tables_data[0]['data']
        for i, row in enumerate(first_table[:10]):  # 처음 10행만
            print(f"행 {i}: {row}")
    
    # 동별로 데이터 그룹화
    dong_data = group_data_by_dong(tables_data)
    
    if dong_data:
        print(f"\n동별 데이터 그룹화 완료: {len(dong_data)}개 동")
        
        # 엑셀로 저장
        output_file = Path('/Users/seongjaehyeon/project/test123/동탄더레이크팰리스_동별분석.xlsx')
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for dong_name in sorted(dong_data.keys(), key=lambda x: int(re.search(r'\d+', x).group())):
                rows = dong_data[dong_name]
                
                # 최대 열 개수 찾기
                max_cols = max(len(row) for row in rows)
                
                # DataFrame 생성
                df = pd.DataFrame(rows)
                
                # 시트 이름 (최대 31자)
                sheet_name = dong_name[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
                print(f"  {dong_name}: {len(rows)}개 행")
        
        print(f"\n✅ 엑셀 파일이 생성되었습니다: {output_file}")
    else:
        print("\n동 데이터를 찾을 수 없습니다.")
        print("표 데이터를 직접 엑셀로 저장합니다.")
        
        # 모든 표를 하나의 엑셀 파일로
        if tables_data:
            output_file = Path('/Users/seongjaehyeon/project/test123/hwpx_all_tables.xlsx')
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for i, table_info in enumerate(tables_data[:50]):  # 최대 50개 표
                    df = pd.DataFrame(table_info['data'])
                    sheet_name = f"표{i+1}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            
            print(f"\n✅ 모든 표를 엑셀로 저장했습니다: {output_file}")

if __name__ == "__main__":
    main()
