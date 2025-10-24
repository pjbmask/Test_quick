import win32com.client
import pandas as pd
from pathlib import Path
import logging
from datetime import datetime
import os

# 로그 설정
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
log_file = log_dir / f"email_send_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class OutlookEmailSender:
    def __init__(self, recipients_file):
        """
        아웃룩 이메일 발송 클래스
        
        Args:
            recipients_file: 수신자 정보가 담긴 엑셀 파일 경로
        """
        self.recipients_file = recipients_file
        self.outlook = None
        self.results = []
        
    def initialize_outlook(self):
        """Outlook 애플리케이션 초기화"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            logging.info("Outlook 초기화 성공")
            return True
        except Exception as e:
            logging.error(f"Outlook 초기화 실패: {e}")
            return False
    
    def load_recipients(self):
        """
        수신자 정보 엑셀 파일 로드
        
        Returns:
            DataFrame: 수신자 정보
        """
        try:
            df = pd.read_excel(self.recipients_file)
            logging.info(f"수신자 정보 로드 완료: {len(df)}명")
            return df
        except Exception as e:
            logging.error(f"수신자 정보 로드 실패: {e}")
            return None
    
    def validate_files(self, file_paths):
        """
        첨부파일 존재 여부 확인
        
        Args:
            file_paths: 파일 경로 리스트
            
        Returns:
            tuple: (유효한 파일 리스트, 누락된 파일 리스트)
        """
        valid_files = []
        missing_files = []
        
        for file_path in file_paths:
            if pd.isna(file_path) or str(file_path).strip() == '':
                continue
            
            file_path = str(file_path).strip()
            
            # 경로 구분자가 없으면 attachments 폴더로 가정
            if not ('/' in file_path or '\\' in file_path):
                file_path = f"attachments/{file_path}"
            
            if Path(file_path).exists():
                valid_files.append(file_path)
            else:
                missing_files.append(file_path)
        
        return valid_files, missing_files
    
    def create_email_body(self, name, position, file_count):
        """
        이메일 본문 생성
        
        Args:
            name: 담당자명
            position: 직급
            file_count: 첨부파일 개수
            
        Returns:
            str: 이메일 본문
        """
        body = f"""
{name} {position}님, 안녕하세요.

내부통제팀입니다.

{name} {position}님께서 담당하시는 영역의 내부통제 관련 자료를 첨부하여 송부드립니다.

※ 첨부파일: {file_count}건

첨부파일 확인 후 문의사항이 있으시면 언제든지 연락 부탁드립니다.

감사합니다.

---
내부통제팀
"""
        return body
    
    def send_email(self, recipient_email, name, position, subject, body, attachments):
        """
        개별 이메일 발송
        
        Args:
            recipient_email: 수신자 이메일
            name: 담당자명
            position: 직급
            subject: 이메일 제목
            body: 이메일 본문
            attachments: 첨부파일 리스트
            
        Returns:
            dict: 발송 결과
        """
        result = {
            '담당자명': name,
            '직급': position,
            '이메일': recipient_email,
            '첨부파일수': len(attachments),
            '상태': '',
            '메시지': ''
        }
        
        try:
            # 이메일 객체 생성
            mail = self.outlook.CreateItem(0)  # 0: olMailItem
            mail.To = recipient_email
            mail.Subject = subject
            mail.Body = body
            
            # 첨부파일 추가
            for file_path in attachments:
                mail.Attachments.Add(str(Path(file_path).absolute()))
            
            # 발송 (또는 임시 보관함에 저장)
            mail.Send()  # 즉시 발송
            # mail.Save()  # 임시 보관함에 저장 (발송 전 확인용)
            
            result['상태'] = '성공'
            result['메시지'] = f'{len(attachments)}개 파일 첨부하여 발송 완료'
            logging.info(f"✓ {name} {position}님 ({recipient_email}) - 발송 성공")
            
        except Exception as e:
            result['상태'] = '실패'
            result['메시지'] = str(e)
            logging.error(f"✗ {name} {position}님 ({recipient_email}) - 발송 실패: {e}")
        
        return result
    
    def process_all(self, subject_template="[내부통제] {date} 담당자별 보고서 송부"):
        """
        전체 이메일 발송 프로세스 실행
        
        Args:
            subject_template: 이메일 제목 템플릿
        """
        logging.info("=" * 60)
        logging.info("이메일 자동 발송 시작")
        logging.info("=" * 60)
        
        # Outlook 초기화
        if not self.initialize_outlook():
            return
        
        # 수신자 정보 로드
        df = self.load_recipients()
        if df is None:
            return
        
        # 제목 생성 (날짜 포함)
        today = datetime.now().strftime('%Y.%m.%d')
        subject = subject_template.format(date=today)
        
        # 각 수신자별로 이메일 발송
        for idx, row in df.iterrows():
            logging.info(f"\n[{idx+1}/{len(df)}] 처리 중...")
            
            name = row['담당자명']
            position = row['직급']
            email = row['이메일']
            
            # 첨부파일 수집 (파일1 ~ 파일10 컬럼)
            file_columns = [col for col in df.columns if col.startswith('파일')]
            file_paths = [row[col] for col in file_columns]
            
            # 파일 유효성 검사
            valid_files, missing_files = self.validate_files(file_paths)
            
            if missing_files:
                logging.warning(f"  누락된 파일: {missing_files}")
            
            if not valid_files:
                result = {
                    '담당자명': name,
                    '직급': position,
                    '이메일': email,
                    '첨부파일수': 0,
                    '상태': '실패',
                    '메시지': '첨부할 유효한 파일이 없음'
                }
                self.results.append(result)
                logging.error(f"  {name} {position}님 - 첨부파일 없음, 발송 스킵")
                continue
            
            # 이메일 본문 생성
            body = self.create_email_body(name, position, len(valid_files))
            
            # 이메일 발송
            result = self.send_email(email, name, position, subject, body, valid_files)
            self.results.append(result)
        
        # 결과 저장
        self.save_results()
        
        logging.info("\n" + "=" * 60)
        logging.info("이메일 발송 완료")
        logging.info("=" * 60)
    
    def save_results(self):
        """발송 결과를 엑셀 파일로 저장"""
        if not self.results:
            return
        
        results_df = pd.DataFrame(self.results)
        result_file = log_dir / f"발송결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        results_df.to_excel(result_file, index=False, engine='openpyxl')
        
        # 결과 요약
        success_count = len(results_df[results_df['상태'] == '성공'])
        fail_count = len(results_df[results_df['상태'] == '실패'])
        
        logging.info(f"\n📊 발송 결과 요약:")
        logging.info(f"  - 성공: {success_count}건")
        logging.info(f"  - 실패: {fail_count}건")
        logging.info(f"  - 결과 파일: {result_file}")


def main():
    """메인 실행 함수"""
    # 수신자 정보 파일 경로
    recipients_file = "reference/recipients.xlsx"
    
    # 파일 존재 확인
    if not Path(recipients_file).exists():
        print(f"❌ 오류: '{recipients_file}' 파일을 찾을 수 없습니다.")
        print("reference 폴더에 수신자 정보 엑셀 파일을 생성해주세요.")
        return
    
    # 이메일 발송 실행
    sender = OutlookEmailSender(recipients_file)
    sender.process_all()
    
    print("\n✅ 프로그램 실행 완료!")
    print(f"📁 로그 확인: {log_file}")


if __name__ == "__main__":
    main()
