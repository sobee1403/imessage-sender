on formatPhoneNumber(phoneNumber)
    # 모든 공백과 특수문자 제거
    set cleanNumber to do shell script "echo " & quoted form of phoneNumber & " | sed 's/[^0-9]//g'"
    
    # 한국 전화번호 형식 확인 및 변환
    if (length of cleanNumber is 11) and (text 1 of cleanNumber is "0") then
        return "+82" & text 2 thru -1 of cleanNumber
    else if (length of cleanNumber is 10) and (text 1 of cleanNumber is "0") then
        return "+82" & text 2 thru -1 of cleanNumber
    else if text 1 thru 2 of cleanNumber is "82" then
        return "+" & cleanNumber
    else
        return phoneNumber
    end if
end formatPhoneNumber

on sendMessage(phoneNumber, messageText)
    set formattedPhone to formatPhoneNumber(phoneNumber)
    
    tell application "Messages"
        set targetService to 1st service whose service type = iMessage
        set targetBuddy to buddy formattedPhone of targetService
        send messageText to targetBuddy
        delay 2
    end tell
end sendMessage

on run
    # Python 스크립트 실행하여 최신 데이터 가져오기
    do shell script "python3 fetch_contacts.py"
    
    # CSV 파일 읽기
    set csvFile to (path to desktop as text) & "message_list.csv"
    set csvData to do shell script "cat " & quoted form of POSIX path of csvFile
    
    # 각 줄을 처리
    set AppleScript's text item delimiters to return
    set rows to text items of csvData
    
    # 첫 번째 줄(헤더) 건너뛰기
    repeat with i from 2 to count of rows
        set currentRow to item i of rows
        
        # 쉼표로 분리된 데이터 처리
        set AppleScript's text item delimiters to ","
        set rowData to text items of currentRow
        
        if (count of rowData) ≥ 2 then
            set phoneNumber to item 3 of rowData  # C열의 전화번호
            set nameText to item 2 of rowData     # B열의 이름
            set messageText to "안녕하세요, " & nameText & "님"  # 원하는 메시지 형식으로 수정
            
            # 메시지 전송
            sendMessage(phoneNumber, messageText)
        end if
    end repeat
end run 