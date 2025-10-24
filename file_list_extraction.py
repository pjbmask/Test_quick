import os
import pandas as pd
from pathlib import Path
from datetime import datetime

def extract_file_list(folder_path, output_excel=None):
    """
    특정 폴더 내 모든 파일 목록을 추출하여 엑셀로 저장
    
    Parameters:
    folder_path (str): 검색할 폴더 경로
    output_excel (str): 출력할 엑셀 파일명 (None이면 자동 생성)
    """
    
    # 파일명에 년월일시분초 추가
    if output_excel is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel = f'파일목록_{timestamp}.xlsx'
    
    # 파일 정보를 저장할 리스트
    file_data = []
    
    # 폴더 내 모든 파일 탐색 (하위 폴더 포함)
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # 절대 경로
            absolute_path = os.path.abspath(os.path.join(root, file))
            
            # 파일명
            file_name = file
            
            # 확장자 (확장자가 없으면 빈 문자열)
            extension = os.path.splitext(file)[1]
            
            file_data.append({
                'A_절대경로': absolute_path,
                'B_파일명': file_name,
                'C_확장자': extension
            })
    
    # 데이터프레임 생성
    df = pd.DataFrame(file_data)
    
    # 엑셀 파일로 저장
    df.to_excel(output_excel, index=False, engine='openpyxl')
    
    print(f"총 {len(file_data)}개의 파일이 추출되었습니다.")
    print(f"결과 파일: {output_excel}")
    
    return df

# 사용 예시
if __name__ == "__main__":
    # 검색할 폴더 경로 지정
    folder_path = r"D:\1. IAM업무\★ 내부회계업무\★ 통제\3. PLC(Process Level Control)\※연도별 평가\19. 25년 운영평가\설계운영평가서"
    
    # 함수 실행 (파일명 자동 생성)
    df = extract_file_list(folder_path)
    
    # 결과 미리보기
    print("\n[추출된 파일 목록 미리보기]")
    print(df.head(10))
