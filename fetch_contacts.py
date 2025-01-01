from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle
from datetime import datetime
import subprocess
from config import SPREADSHEET_ID

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def get_current_week_sheet_name():
    now = datetime.now()
    month = now.month
    week_of_month = (now.day - 1) // 7 + 1
    return f"{month}월 {week_of_month}주차"

def get_google_sheet_data():
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

        # 디버깅을 위한 출력
        print("가져온 데이터:")
        print("사원정보:", employee_data)
        print("보고데이터:", report_data)

        # iMessage 대상 찾기
        messages_to_send = []
        for row in report_data:
            if len(row) >= 2 and (len(row[0].strip()) == 0 or row[0].strip() == ''):
                employee_name = row[1]
                for emp in employee_data:
                    if len(emp) >= 5 and emp[0] == employee_name and emp[4] == 'imessage':
                        messages_to_send.append([emp[0], emp[1]])  # [이름, 전화번호]

        return messages_to_send

    except Exception as e:
        print(f'에러가 발생했습니다: {e}')
        return []

def send_imessage(phone_number, message):
    # 전화번호가 비어있거나 유효하지 않은 경우 처리
    if not phone_number or len(phone_number.strip()) == 0:
        print(f"전화번호가 비어있습니다.")
        return False

    # 전화번호 형식 정리 (하이픈 제거 등)
    phone_number = phone_number.replace("-", "").replace(" ", "")
    
    applescript = f'''
    tell application "Messages"
        set targetBuddy to "{phone_number}"
        set targetService to id of 1st service whose service type = iMessage
        set theBuddy to buddy targetBuddy of service id targetService
        send "{message}" to theBuddy
    end tell
    '''
    
    try:
        subprocess.run(['osascript', '-e', applescript], check=True)
        print(f"메시지 전송 성공: {phone_number}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"메시지 전송 실패: {phone_number}, 에러: {e}")
        return False

def get_weekly_message(employee_name):
    current_sheet = get_current_week_sheet_name()
    return f"안녕하세요, {employee_name}님.{current_sheet} 주간업무보고가 미제출되었습니다. 확인 부탁드립니다."

if __name__ == '__main__':
    messages = get_google_sheet_data()
    print("iMessage 전송을 시작합니다...")
    
    for msg in messages:
        employee_name = msg[0]
        phone = msg[1]
        
        # 전화번호 유효성 검사
        if not phone or len(phone.strip()) == 0:
            print(f"✗ {employee_name}님의 전화번호가 없습니다. 건너뜁니다.")
            continue
            
        message = get_weekly_message(employee_name)
        
        print(f"\n{employee_name}님에게 메시지 전송 중...")
        success = send_imessage(phone, message)
        
        if success:
            print(f"✓ {employee_name}님에게 전송 완료")
        else:
            print(f"✗ {employee_name}님에게 전송 실패") 