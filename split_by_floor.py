import pandas as pd
from pathlib import Path
import json

def split_by_floor_ranges(input_excel, config_file, output_excel):
    """
    엑셀 파일의 각 동 시트를 층별로 나누기
    
    Args:
        input_excel: 입력 엑셀 파일 경로 (동별로 나뉜 파일)
        config_file: 층별 범위 설정 JSON 파일
        output_excel: 출력 엑셀 파일 경로
    """
    
    # 설정 파일 읽기
    with open(config_file, 'r', encoding='utf-8') as f:
        floor_config = json.load(f)
    
    print("=== 층별 범위 설정 ===")
    for dong, floors in floor_config.items():
        print(f"\n{dong}:")
        for floor_name, range_info in floors.items():
            print(f"  {floor_name}: {range_info['start_floor']}층 ~ {range_info['end_floor']}층")
    
    # 입력 엑셀 파일 읽기
    excel_file = pd.ExcelFile(input_excel)
    
    print(f"\n입력 파일의 시트: {excel_file.sheet_names}")
    
    # 출력 엑셀 파일 생성
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        
        # 각 동(시트)별로 처리
        for sheet_name in excel_file.sheet_names:
            print(f"\n처리 중: {sheet_name}")
            
            # 시트 데이터 읽기
            df = pd.read_excel(input_excel, sheet_name=sheet_name, header=None)
            
            print(f"  전체 행 수: {len(df)}")
            
            # 해당 동의 설정이 있는지 확인
            if sheet_name in floor_config:
                floors = floor_config[sheet_name]
                
                # 하나의 시트에 모든 층 데이터를 결합
                combined_data = []
                
                # 각 층별로 데이터 추출하여 결합
                for floor_name, range_info in floors.items():
                    start_floor = range_info['start_floor']
                    end_floor = range_info['end_floor']
                    
                    # 층 제목 행 추가
                    title_row = [f"[ {floor_name} ]"] + [''] * (df.shape[1] - 1)
                    combined_data.append(title_row)
                    
                    # 해당 층 범위에 속하는 행 찾기
                    floor_rows = []
                    for idx, row in df.iterrows():
                        row_text = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                        
                        # 층 패턴 찾기 (예: "15층", "옥탑 2층", "지붕층" 등)
                        floor_patterns = []
                        
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
                    if floor_rows:
                        combined_data.extend(floor_rows)
                        print(f"  ✓ {floor_name}: {len(floor_rows)}개 행")
                    else:
                        print(f"  ⚠ {floor_name}: 데이터 없음")
                    
                    # 구분을 위한 빈 행 추가
                    empty_row = [''] * df.shape[1]
                    combined_data.append(empty_row)
                
                # 결합된 데이터를 DataFrame으로 변환
                combined_df = pd.DataFrame(combined_data)
                
                # 엑셀에 저장
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                print(f"  → '{sheet_name}' 시트에 {len(floors)}개 층 그룹 통합")
            else:
                # 설정이 없는 동은 그대로 복사
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                print(f"  ⚠ 설정 없음. 원본 그대로 복사")
    
    print(f"\n✅ 층별 분리 완료: {output_excel}")

def main():
    # 파일 경로 설정
    input_file = Path('/Users/seongjaehyeon/project/test123/동탄더레이크팰리스_동별분석.xlsx')
    config_file = Path('/Users/seongjaehyeon/project/test123/floor_config.json')
    output_file = Path('/Users/seongjaehyeon/project/test123/동탄더레이크팰리스_층별분석.xlsx')
    
    if not input_file.exists():
        print(f"❌ 입력 파일을 찾을 수 없습니다: {input_file}")
        return
    
    if not config_file.exists():
        print(f"❌ 설정 파일을 찾을 수 없습니다: {config_file}")
        print("\n💡 floor_config.json 파일을 먼저 생성하고 층별 범위를 설정하세요.")
        print("\n예시:")
        print('''{
  "2561동": {
    "저층부(1~8층)": {"start_floor": 1, "end_floor": 8},
    "중층부(9~15층)": {"start_floor": 9, "end_floor": 15}
  }
}''')
        return
    
    # 층별로 나누기
    split_by_floor_ranges(input_file, config_file, output_file)

if __name__ == "__main__":
    main()
