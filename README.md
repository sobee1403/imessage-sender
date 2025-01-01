# iMessage Sender ver.20241224.01

Google Sheets에서 데이터를 읽어와 미제출자에게 자동으로 iMessage를 보내는 Python 스크립트입니다.

## 기능
- Google Sheets API를 사용하여 스프레드시트 데이터 읽기
- 주간 업무보고 미제출자 자동 확인
- iMessage를 통한 자동 알림 발송

## 필요 조건
- Python 3.x
- macOS (iMessage 기능 사용을 위해)
- Google Cloud Console 프로젝트 및 인증 설정
- 활성화된 iMessage 계정

## 설치 방법
1. 저장소 클론
bash
git clone https://github.com/[username]/imessage-sender.git
cd imessage-sender

2. 가상환경 생성 및 활성화

bash
git clone https://github.com/[username]/imessage-sender.git
cd imessage-sender
활성화
bash
python3 -m venv venv
source venv/bin/activate

3. 필요한 패키지 설치

bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib

## 사용 방법
1. Google Cloud Console에서 credentials.json 파일을 다운로드하여 프로젝트 루트 디렉토리에 저장

2. 스크립트 실행
bash
python3 fetch_contacts.py

## 스프레드시트 구조
1. 사원정보 시트
   - B열: 사원 이름
   - C열: 휴대폰 번호
   - D열: 이메일 주소
   - E열: 카카오 아이디
   - F열: 알림 방식 (현재는 'imessage'만 지원)

2. 주간업무보고 시트 (시트명 예: "1월 1주차")
   - B3:C11 범위에서 데이터 읽기
   - B열이 비어있는 경우 미제출로 간주
   - C열: 사원명

## 주의사항
- credentials.json 파일은 보안을 위해 절대 공유하지 마세요
- 개인정보가 포함된 token.pickle 파일도 공유하지 마세요
- macOS의 Messages 앱에 로그인되어 있어야 합니다

## 라이선스
이 프로젝트는 MIT 라이선스를 따릅니다.




