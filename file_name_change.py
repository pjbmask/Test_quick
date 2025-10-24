import pandas as pd
import os
import shutil

# 엑셀 파일 읽기
excel_file = r"C:\py\Test_quick\reference\rename.xlsx"  # 엑셀 파일 경로를 지정하세요
df = pd.read_excel(excel_file)

# A열: 원본 파일 경로, B열: 변경할 파일 경로
for index, row in df.iterrows():
    old_path = row['Before']  # A열의 열 이름에 맞게 수정하세요
    new_path = row['After']  # B열의 열 이름에 맞게 수정하세요
    
    # 파일 존재 여부 확인
    if os.path.exists(old_path):
        try:
            # 새 경로의 디렉토리가 없으면 생성
            new_dir = os.path.dirname(new_path)
            if new_dir and not os.path.exists(new_dir):
                os.makedirs(new_dir)
            
            # 파일 이름 변경 (이동)
            shutil.move(old_path, new_path)
            print(f"성공: {old_path} -> {new_path}")
            
        except Exception as e:
            print(f"실패: {old_path} - 오류: {str(e)}")
    else:
        print(f"파일 없음: {old_path}")

print("\n파일명 변경 작업 완료")
