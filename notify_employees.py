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

def get_google_sheet_data():
    """미제출자 정보를 가져옵니다."""
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

        # 1. 사원정보 시트에서 데이터 가져오기
        employee_range = '사원정보!B2:F9'
        employee_result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=employee_range
        ).execute()
        employee_data = employee_result.get('values', [])

        # 2. 주간업무보고 시트에서 데이터 가져오기
        current_sheet = get_current_week_sheet_name()
        report_range = f'{current_sheet}!B3:C11'
        report_result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=report_range
        ).execute()
        report_data = report_result.get('values', [])

        # iMessage 대상 찾기
        messages_to_send = []
        for row in report_data:
            if len(row) >= 2 and (len(row[0].strip()) == 0 or row[0].strip() == ''):
                employee_name = row[1]
                # 사원정보에서 해당 사원의 정보 찾기
                for emp in employee_data:
                    if emp[0] == employee_name and len(emp) >= 4 and emp[4] == 'imessage':
                        messages_to_send.append({
                            'name': emp[0],
                            'phone': emp[1] if len(emp[1].strip()) > 0 else '',
                            'email': emp[2] if len(emp) > 2 and len(emp[2].strip()) > 0 else '',
                        })

        return messages_to_send

    except Exception as e:
        print(f'에러가 발생했습니다: {e}')
        return []

def get_weekly_message(employee_name):
    """미제출자에게 보낼 메시지를 생성합니다."""
    current_sheet = get_current_week_sheet_name()
    return f"안녕하세요, {employee_name}님. {current_sheet} 주간업무보고가 미제출되었습니다. 확인 부탁드립니다.\n\n@https://url.kr/7dmp8n"

if __name__ == '__main__':
    messages = get_google_sheet_data()
    print("미제출자 알림 메시지 전송을 시작합니다...")
    
    for msg in messages:
        employee_name = msg['name']
        phone = msg['phone']
        email = msg['email']
        
        contact = phone if phone else email
        if not contact:
            print(f"✗ {employee_name}님의 연락처 정보가 없습니다.")
            continue
            
        message = get_weekly_message(employee_name)
        print(f"\n{employee_name}님에게 메시지 전송 중...")
        success = send_imessage(contact, message)
        
        if success:
            print(f"✓ {employee_name}님에게 전송 완료")
        else:
            print(f"✗ {employee_name}님에게 전송 실패") 