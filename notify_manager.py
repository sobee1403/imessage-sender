from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle
from datetime import datetime
from config import SPREADSHEET_ID
from imessage_sender import send_imessage, get_current_week_sheet_name

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def get_managers_contacts():
    """사원정보 시트에서 관리자 정보를 가져옵니다."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        # 사원정보 시트에서 관리자 정보 가져오기 (B:G 열)
        range_name = '사원정보!B2:G9'
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        values = result.get('values', [])
        
        managers = []
        for row in values:
            if len(row) >= 6 and row[5] == '관리자':  # G열이 '관리자'인 경우
                managers.append({
                    'name': row[0],  # 이름
                    'phone': row[1],  # 전화번호
                    'email': row[2]   # 이메일
                })
        return managers
    
    except Exception as e:
        print(f'관리자 정보 조회 중 에러 발생: {e}')
        return []

def get_submission_status():
    """제출 현황을 조회합니다."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        
        current_sheet = get_current_week_sheet_name()
        report_range = f'{current_sheet}!B3:C11'
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=report_range
        ).execute()
        values = result.get('values', [])
        
        total_count = len(values)
        submitted = sum(1 for row in values if row and len(row) > 0 and len(row[0].strip()) > 0)
        not_submitted = total_count - submitted
        not_submitted_names = [row[1] for row in values if len(row) > 1 and (len(row[0].strip()) == 0)]
        
        return {
            'total': total_count,
            'submitted': submitted,
            'not_submitted': not_submitted,
            'not_submitted_names': not_submitted_names
        }
    except Exception as e:
        print(f'제출 현황 조회 중 에러 발생: {e}')
        return None

def get_manager_message(status):
    """관리자에게 보낼 메시지를 생성합니다."""
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
    current_sheet = get_current_week_sheet_name()
    
    message = f"[{current_sheet} 제출 현황 보고 - {current_time}]\n"
    message += f"총 인원: {status['total']}명\n"
    message += f"제출 완료: {status['submitted']}명\n"
    message += f"미제출: {status['not_submitted']}명\n"
    if status['not_submitted_names']:
        message += f"미제출자: {', '.join(status['not_submitted_names'])}\n"
    message += "\n@https://url.kr/7dmp8n"
    
    return message

if __name__ == '__main__':
    print("관리자 보고 메시지 전송을 시작합니다...")
    
    # 관리자 목록 조회
    managers = get_managers_contacts()
    if not managers:
        print("관리자 정보를 찾을 수 없습니다.")
        exit(1)
        
    # 제출 현황 조회
    status = get_submission_status()
    if not status:
        print("제출 현황 조회에 실패했습니다.")
        exit(1)
        
    # 각 관리자에게 메시지 전송
    message = get_manager_message(status)
    for manager in managers:
        print(f"\n{manager['name']} 관리자에게 보고 전송 중...")
        contact = manager['phone'] if manager['phone'] else manager['email']
        
        if not contact:
            print(f"✗ {manager['name']} 관리자의 연락처 정보가 없습니다.")
            continue
            
        success = send_imessage(contact, message)
        if success:
            print(f"✓ {manager['name']} 관리자에게 보고 완료")
        else:
            print(f"✗ {manager['name']} 관리자에게 보고 실패") 