import subprocess
from datetime import datetime

def get_current_week_sheet_name():
    now = datetime.now()
    month = now.month
    week_of_month = (now.day - 1) // 7 + 1
    return f"{month}월 {week_of_month}주차"

def send_imessage(contact, message):
    """
    전화번호나 이메일로 iMessage 발송
    Args:
        contact (str): 수신자의 전화번호 또는 이메일
        message (str): 발송할 메시지
    """
    if not contact or len(contact.strip()) == 0:
        print(f"연락처 정보가 없습니다.")
        return False

    if contact.replace('-', '').isdigit():
        contact = contact.replace("-", "").replace(" ", "")
    
    applescript = f'''
    tell application "Messages"
        set targetBuddy to "{contact}"
        set targetService to id of 1st service whose service type = iMessage
        set theBuddy to buddy targetBuddy of service id targetService
        send "{message}" to theBuddy
    end tell
    '''
    
    try:
        subprocess.run(['osascript', '-e', applescript], check=True)
        print(f"메시지 전송 성공: {contact}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"메시지 전송 실패: {contact}, 에러: {e}")
        return False

if __name__ == '__main__':
    # 테스트용 코드
    test_message = "테스트 메시지입니다."
    test_contact = "your-test-contact"  # 테스트용 연락처
    send_imessage(test_contact, test_message) 