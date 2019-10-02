import telnetlib
import time
import win32com.client
#엑셀 프로그램 실행
excel = win32com.client.Dispatch("Excel.Application")
#엑셀 프로그램 보이게 하는 설정
excel.Visible = True
#엑셀 경로 r을 붙이면 정규표현식이 된다고 함
Path = r'D:\python\autobackup\ping_test.xlsx'
#해당 경로의 엑셀 파일 실행
wb = excel.Workbooks.Open(Path)
#실행된 엑셀 파일에 가장 먼저 열린 sheet
ws = wb.ActiveSheet
#while문 쓰려고 변수 1을 줌
r=1
#while 반복문 엑셀 데이터가 없을 때 까지 반복
while True:
    #엑셀 데이터 아래로 한칸씩 이동
    r = r + 1
    #엑셀의 셀값이 없을 때(none) 끝내기 None이 어떤 형인지 모르겠음. 일단 str 문자열로 변형하여 None로 비교
    if str(ws.Cells(r,1))=="None":
        print("나가라")
        break
    print("r의값",r)
    #엑셀에서 IP를 가져오는 좌표
    HOST=ws.Cells(r,1).Value
    #엑셀에서 port를 가져오는 좌표 엑셀에서 가져올때 23.0 으로 소수점을 가져와서 int로 소수점을 지움
    PORT=int(ws.Cells(r,3).Value)
    #엑셀에서 백업된 config 파일 이름(고객사 이름)
    CONFIG_name=ws.Cells(r,4).Value
    #telnet ID
    TACACS_ID = b"sysone"
    #telnet 비밀번호
    TACACS_PW = b"!tltmdnjs2012@"
    #ftp백업 명령어
    cmd = b"execute backup config ftp "
    #인코딩 디코딩 어려워서 명령어를 3등분으로 나눔
    cmd1 = CONFIG_name.encode('euc-kr')#엑셀에서 값을 불러올 때 유니코드로 가져오는듯 그걸 EUC-kr로 인코딩
    cmd2 = b".conf 192.168.201.2 anonymous anonymous"
    print(cmd)
    #login 이라는 문자를 기다릴거임. 근데 telnet 모듈에서 바이트형을 받아야 해서 ASCII로 인코딩함
    Login_prompt = "login:".encode('ASCII')
    #예외처리를 위해 try문 사용 예외처리하면 중간에 에러 나도 무시하고 진행
    try:
        print(PORT)
        #telnet 모듈 이용하여 telnet 접속 IP,port
        ko_telnet = telnetlib.Telnet(HOST,PORT)
        print(ko_telnet)
        print("확인\n")
        #telnet에서 원하는 문자열이 나올때 까지 기다리는 함수 (기다리는 문자, 기다리는 시간->시간 오버되면 넘어감)
        response = ko_telnet.read_until(Login_prompt, 3)
        #telnet에 텍스트 넣기 ID
        ko_telnet.write(TACACS_ID + b"\n")
        #telnet에 ID를 넣고 비밀번호를 넣기위해 Password 문자 기다림
        response = ko_telnet.read_until("Password:".encode('ASCII'), 3)
        #telnet에 텍스트 넣기 비밀번호
        ko_telnet.write(TACACS_PW + b"\n")
        #로그인 지연 있을 수 있으니 3초 지연
        time.sleep(3)
        #로그인 성공 시 Welcome 문자 검색
        response = ko_telnet.read_until("Welcome !".encode('ASCII'), 3)
        #if문으로 성공 실패 동작 구분
        if b"Welcome" in response:
            #일단 로그인 성공 하면 엑셀에 connect로 기재
            ws.Cells(r, 2).Value = 'connect'
            print(cmd+cmd1+cmd2+b"\n")
            #3등분으로 나눈 명령어 telent으로 전달
            ko_telnet.write(cmd+cmd1+cmd2+b"\n")
            #config 백업 완료 시 발생하는 문자 대기_잘안됨....
            response1 = ko_telnet.read_until("Send config file to ftp server OK.".encode('ASCII'), 20)
            print("response1 값")
            print(response1)
            print("값")
            #if문으로 백업 성공 유무 확인
            if b"server OK." in response1:
                #성공 시 엑셀에 OK기재
                ws.Cells(r, 5).Value = 'OK'
            else :
                #실패 시 엑셀에 fail기재
                ws.Cells(r, 5).Value = 'fail'
            #백업 후 exit로 telnet 종료
            ko_telnet.write(b"exit" + b"\n")
            print(ko_telnet.read_all())
        else:
            #telent접근 실패 시 엑셀에no connect 기재
            ws.Cells(r, 2).Value = 'no connect'
            print(response)
            #계정정보 틀림으로 접근 안될 때 3번 시도함.
            ko_telnet.write(TACACS_ID + b"\n")
            time.sleep(1)
            ko_telnet.write(TACACS_PW + b"\n")
            time.sleep(1)
            ko_telnet.write(TACACS_ID + b"\n")
            time.sleep(1)
            ko_telnet.write(TACACS_PW + b"\n")
            time.sleep(1)
            print("else 마지막")
    except:
        #예외 처리  목적 telnet IP가 달라서 접근 자체가 안되서 telnet 모듈이 종료되는것을 방지
        print("접근안됨")
        ws.Cells(r, 2).Value = 'no Host'
        pass



print("끝")
#엑셀 파일 저장
wb.Save()
#엑셀 종료
excel.Quit()



