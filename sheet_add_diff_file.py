import win32com.client as win32
import openpyxl
import os

def copy_sheet_to_files(file_list_excel):
    """
    B열 템플릿 파일의 시트를 A열 대상 파일의 마지막에 추가 (Excel COM 사용)
    """
    
    # 1. 파일 목록 읽기
    print("파일 목록을 읽는 중...")
    list_wb = openpyxl.load_workbook(file_list_excel)
    list_ws = list_wb.active
    
    file_pairs = []
    for row in range(2, list_ws.max_row + 1):
        target_path = list_ws[f'A{row}'].value
        template_path = list_ws[f'B{row}'].value
        
        if target_path and template_path:
            target_path = str(target_path)
            template_path = str(template_path)
            
            # 절대 경로로 변환
            target_path = os.path.abspath(target_path)
            template_path = os.path.abspath(template_path)
            
            if not os.path.exists(target_path):
                print(f"경고: 대상 파일을 찾을 수 없음 - {target_path}")
                continue
            if not os.path.exists(template_path):
                print(f"경고: 템플릿 파일을 찾을 수 없음 - {template_path}")
                continue
                
            file_pairs.append((target_path, template_path))
    
    print(f"총 {len(file_pairs)}개 작업 발견\n")
    list_wb.close()
    
    # 2. Excel 실행
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # 백그라운드 실행
    excel.DisplayAlerts = False  # 경고창 끄기
    
    success_count = 0
    fail_count = 0
    
    try:
        for idx, (target_file, template_file) in enumerate(file_pairs, 1):
            try:
                print(f"[{idx}/{len(file_pairs)}]")
                print(f"  대상: {os.path.basename(target_file)}")
                print(f"  템플릿: {os.path.basename(template_file)}")
                
                # 파일 열기
                template_wb = excel.Workbooks.Open(template_file)
                target_wb = excel.Workbooks.Open(target_file)
                
                # 템플릿의 첫 번째 시트
                template_sheet = template_wb.Worksheets(1)
                template_sheet_name = template_sheet.Name
                
                # 대상 파일에 동일 이름 시트가 있으면 삭제
                for sheet in target_wb.Worksheets:
                    if sheet.Name == template_sheet_name:
                        print(f"  - 기존 '{template_sheet_name}' 시트 삭제")
                        sheet.Delete()
                        break
                
                # 시트를 대상 파일의 마지막으로 복사
                last_sheet = target_wb.Worksheets(target_wb.Worksheets.Count)
                template_sheet.Copy(After=last_sheet)
                
                print(f"  - '{template_sheet_name}' 시트 추가 완료")
                
                # 저장 및 닫기
                target_wb.Save()
                target_wb.Close()
                template_wb.Close()
                
                print(f"  ✓ 완료\n")
                success_count += 1
                
            except Exception as e:
                print(f"  ✗ 실패: {str(e)}\n")
                fail_count += 1
                
                # 열린 파일 정리
                try:
                    target_wb.Close(SaveChanges=False)
                except:
                    pass
                try:
                    template_wb.Close(SaveChanges=False)
                except:
                    pass
    
    finally:
        # Excel 종료
        excel.Quit()
    
    # 결과 출력
    print("=" * 50)
    print(f"작업 완료!")
    print(f"성공: {success_count}개")
    print(f"실패: {fail_count}개")
    print("=" * 50)


# 사용 예시
if __name__ == "__main__":
    file_list_excel = r"C:\py\Test_quick\reference\파일목록_결과.xlsx"
    copy_sheet_to_files(file_list_excel=file_list_excel)
