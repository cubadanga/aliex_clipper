import time
import re
import os
import sys
import pyperclip
import keyboard
from colorama import Fore
import numpy as np
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs, urlunparse
import pywinauto
import pygetwindow as gw
import configparser
from urllib.request import urlopen
from urllib.error import HTTPError

def loadPassword(): #우선 'set.ini' 파일에 저장된 패스워드와 웹에 있는 패스워드가 일치하는지 확인한다.
    basedir = os.getcwd()
    ini_dir = os.path.join(basedir,'set.ini')

    # pc set.ini 파일의 저장된 pass워드 읽어오기
    properties = configparser.ConfigParser()
    properties.read(ini_dir)
    
    if 'DEFAULT' in properties and 'userpass' in properties['DEFAULT']:
        password = properties['DEFAULT']['userpass']
        return password
    else:
        print(Fore.RED + "오류 - 'userpass' key not found in set.ini file."+'\n')
        print(Fore.RESET + "엔터를 누르면 종료합니다.")
        aInput = input("")
        sys.exit()

def getPtag(url): # 웹페이지에 적어 놓은 password 텍스트를 크롤링해 추출하는 함수
    try:
        html = urlopen(url)
        
    except HTTPError as e:
        print(Fore.RED + '오류 - 네트워크오류 또는 패스워드url오류'+'\n')
        print(Fore.RESET + "엔터를 누르면 종료합니다.")
        aInput = input("")
        sys.exit()
    try:
        soup = BeautifulSoup(html,"html.parser")
        ptag = soup.find('p')
        
    except AttributeError as e:
        return None
    return ptag.text

def judge(password,passTag): #set.ini에 저장된 패스워드와 웹에 있는 패스워드를 비교하는 함수.
    if password == passTag:
        properties = configparser.ConfigParser()
        properties.set('DEFAULT','userpass',password)
        with open('./set.ini','w',encoding='utf-8') as F:
            properties.write(F)
                
        pass
    else:
        print(Fore.RED + "오류 - 저장된 패스워드가 없거나 올바른 패스워드가 아닙니다."+Fore.RESET+'\n')
        inputPass(password,passTag)

def inputPass(password,passTag): #패스워드가 틀렸을 때 콘솔에서 다시 입력을 받는 함수
    userPass = str(password)
    passTag = passTag
    print('\n' + "패스워드를 입력해 주세요.")
    userPass = input()
    judge(userPass, passTag)

def title_collector(page_min, page_max):
    page_list=[]
    title_list = []   

    for i in range(page_min,page_max+1):
        # Wait for the user to switch to the Chrome window
        time.sleep(2)  # adjust sleep time as needed

        # Execute key combination to focus address bar (Ctrl + L)
        keyboard.press_and_release('ctrl+l')

        # Set low latency
        time.sleep(0.5)

        # Execute key combination to select URL (Ctrl + C)
        keyboard.press_and_release('ctrl+c')
        time.sleep(0.5)
        url = pyperclip.paste()
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        query_params.pop('trafficChannel', None)
        query_params.pop('page', None)
        updated_url = urlunparse(parsed_url._replace(query='&'.join([f"{k}={v[0]}" for k, v in query_params.items()])))

        time.sleep(0.5)
        add_option = f'&trafficChannel=main&g=y&page={i}'
        target_url = updated_url + add_option
        copy_url = pyperclip.copy(target_url)
        keyboard.press_and_release('ctrl+v')
        time.sleep(0.5)

        # Go to URL using keyboard (Enter key)
        keyboard.press_and_release('enter')

            # Wait time for page load         
        time.sleep(3)  # adjust sleep time as needed

        # Specify the number of scroll downs
        num_scrolls = 10

    # Perform the scrolling
        scroll_count = 0
        while scroll_count < num_scrolls:
            keyboard.press_and_release('space')  # Scroll down one page
            time.sleep(0.5)
            scroll_count += 1

    # Wait for the scroll to settle
        # Execute key combination to copy source code (Ctrl + U)
        keyboard.press_and_release('ctrl+u')
        time.sleep(2)
        keyboard.press_and_release('ctrl+a')
        time.sleep(0.5)
        keyboard.press_and_release('ctrl+c')
        time.sleep(0.5)
        keyboard.press_and_release('ctrl+w')

        # Get source code from clipboard
        html = pyperclip.paste()
        soup = BeautifulSoup(html, 'html.parser')
        
        html_txt = str(soup)
        pattern = r'"displayTitle":\s*"(.*?)",'
        display_titles = re.findall(pattern, html_txt)

        if display_titles:
            for title in display_titles:
                clean_title = title.replace(',','')
                title_list.append(clean_title)
        else:
            print("No matches found.")
    result_title = '\n'.join(title_list)
    return result_title
#프로그램 시작
np.set_printoptions(threshold=np.inf, linewidth=np.inf)
print(Fore.LIGHTBLUE_EX + "알리익스프레스 제목 수집기 시작...")
print(Fore.RESET)


password = loadPassword() #set.ini 파일에서 패스워드를 읽는 함수
#개발자 테스트용 페이지
passTag = getPtag("https://sites.google.com/view/testexmaker/home/aliex_title_collector") #관리자 패스워드가 저장된 웹페이지 url을 전달하여 패스워드를 크롤링 해 오는 getPtag 함수 실행
judge(password,passTag) #웹에서 가져온 패스워드와 set.ini 파일에 저장된 패스워드를 비교하여 틀리면 입력창으로 입력받고 맞으면 통과시킴


tday_s = time.strftime('%Y%m%d-%H%M%S',time.localtime(time.time()))

page_num = 1
print('\n' + "몇 페이지부터 수집할까요? 페이지번호 입력")
page_min = int(input())

print('\n' + "몇 페이지까지 수집할까요? 페이지번호 입력")
page_max = int(input())

print(f'\n + {page_min}페이지부터 {page_max}까지 수집합니다!')
print(f'\n + 크롬창에 검색결과는 띄워 놓아야 합니다.')
print(f'\n + 엔터를 누르면 시작!')
ainput = input('')

win = gw.getWindowsWithTitle('AliExpress')[0]

if not win.isActive:
    pywinauto.application.Application().connect(handle=win._hWnd).top_window().set_focus()
    win.activate()
win.maximize()

result_title = title_collector(page_min, page_max)
print(F'추출된 제목 : \n{result_title}')

with open(f"result_data"+'_'+tday_s+'.txt', 'w', encoding="utf-8") as file:
    file.write(result_title)

winTerminal = gw.getWindowsWithTitle('aliex_clipper')[0]

if not winTerminal.isActive:
    pywinauto.application.Application().connect(handle=winTerminal._hWnd).top_window().set_focus()

print("\n*** result_data.txt 파일로 저장완료.")
print("\n엔터를 누르면 종료합니다.")
binput = input('')
