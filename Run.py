from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
import re
import pyperclip
import sys
import os

def Search(s):
    browser.execute_script('selectCategory(0,0, true);')
    Search = browser.find_element_by_id('txtItemName')
    Search.clear()
    Search.send_keys(s)
    Search.send_keys('\n')
    time.sleep(1.5)

def ChangeWindow(n):
    browser.switch_to.window(browser.window_handles[n])

def Login():
    if login_type != "stove":
        browser.find_element_by_id("{}_login".format(login_type)).click()
        time.sleep(1)
        ChangeWindow(1)

    sid = {"stove"  : "user_id", 
           "naver"  : "id",
           "facebook" : "email",
           "twitter" : "username_or_email"}
    spw = {"stove" : "user_pwd",
           "naver" :  "pw",
           "facebook" : "pass",
           "twitter" : "password"}
    sclick = {"stove" : "#idLogin > div.row.grid.el-actions > button",
              "naver" : "#log\.login",
              "facebook" : "#loginbutton",
              "twitter" : "#allow"}

    bid = browser.find_element_by_id(sid[login_type])
    bid.click()
    pyperclip.copy(uid)
    bid.send_keys(Keys.CONTROL, 'v')
    time.sleep(1)

    bpw = browser.find_element_by_id(spw[login_type])
    bpw.click()
    pyperclip.copy(upw)
    bpw.send_keys(Keys.CONTROL, 'v')
    time.sleep(1)
    
    browser.find_element_by_css_selector(sclick[login_type]).click()
    time.sleep(2)
    ChangeWindow(0)

def Message():
    global progress
    progress += 1
    msg = "\r진행률 : %d%%"%(progress)
    print(msg, end='')

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
browser = webdriver.Chrome('chromedriver.exe', options=options)
browser.get('https://lostark.game.onstove.com/Market')

wb = openpyxl.load_workbook('base.xlsm', read_only=False, keep_vba=True)
ws = wb['거래소 최저가']

os.system('cls')
print('\n','-'*43,'\n Lostark Craft Tool ver3.2.0.0 by simm4256\n','-'*43,'\n')

login_type = wb['검색']['I5'].internal_value
uid = wb['검색']['I6'].internal_value
upw = wb['검색']['I7'].internal_value

if login_type==None or uid==None or upw==None:
    browser.quit()
    print("로그인 정보 불러오기에 실패했습니다.")
    print("base.xlsm 파일에 로그인 정보를 잘 입력했는지 확인해주세요.")
    print("만약 Readme.txt를 읽지 않으셨다면 반드시 읽어주세요.")
    print("이 창은 10초 후 닫힙니다.")
    time.sleep(10)
    sys.exit()

try:
    Login()
except:
    print("로그인 과정에서 오류가 생겼습니다.")
    print("STOVE 계정이 아니라면 STOVE 계정으로 시도해보시기 바랍니다.")
    print("이 오류가 반복될 경우 제작자에게 문의해주세요.")
    print("이 창은 10초 후 닫힙니다.")
    time.sleep(10)
    sys.exit()

isError = True
for i in range(5):
    try:
        progress=-1
        err = 1

        Message()
        browser.find_element_by_css_selector('.main-category > li:nth-child(8) > a:nth-child(1)').click(); Message()
        browser.find_element_by_css_selector('.is-active > ul:nth-child(2) > li:nth-child(1) > a:nth-child(1)').click(); Message(); err+=1
        time.sleep(1.5); Message()
        browser.find_element_by_css_selector('#itemList > thead:nth-child(3) > tr:nth-child(1) > th:nth-child(1) > a:nth-child(1)').click(); Message(); err+=1

        cnt=0
        j=1
        for k in range(1,5) :
            if k>1 :
                browser.execute_script('paging.page({});'.format(k))
            err+=1
            time.sleep(0.7); Message(); Message(); Message(); Message(); Message(); Message()
            prices = browser.find_elements_by_class_name("price"); Message()
            names = browser.find_elements_by_class_name('name'); Message()
            cnt_name = 1; Message()
            for i in prices:
                cnt+=1
                if cnt==3 :
                    cnt=0
                    j+=1
                    ws['A{}'.format(j)]=names[cnt_name].text
                    try : 
                        ws['D{}'.format(j)] = re.sub(r'[^0-9]', '', (browser.find_element_by_css_selector('#tbodyItemList > tr:nth-child({}) > td:nth-child(1) > div:nth-child(1) > span:nth-child(3)'.format(cnt_name)).text))
                    except : None
                    cnt_name+=1
                    p=int(i.text.replace(',',""))
                    ws['B{}'.format(j)] = p
            j+=1; err+=1
        browser.find_element_by_css_selector('.main-category > li:nth-child(6) > a:nth-child(1)').click()
        browser.find_element_by_css_selector('.is-active > ul:nth-child(2) > li:nth-child(1) > a:nth-child(1)').click(); err+=1
        time.sleep(1)
        browser.find_element_by_css_selector('#itemList > thead:nth-child(3) > tr:nth-child(1) > th:nth-child(1) > a:nth-child(1)').click(); err+=1

        cnt=0
        j=1
        for k in range(1,7) :
            if k>1 :
                browser.execute_script('paging.page({});'.format(k))
            err+=1
            time.sleep(0.7); Message(); Message(); Message(); Message(); Message(); Message()
            prices = browser.find_elements_by_class_name("price")
            names = browser.find_elements_by_class_name('name')
            cnt_name = 1
            for i in prices:
                cnt+=1
                if cnt==3 :
                    cnt=0
                    j+=1
                    ws['F{}'.format(j)]=names[cnt_name].text
                    cnt_name+=1
                    p=int(i.text.replace(',',""))
                    ws['G{}'.format(j)] = p
            j+=1; err+=1

        Search('하 융'); Message(); Message(); Message(); err+=1
        browser.find_element_by_css_selector('#itemList > thead:nth-child(3) > tr:nth-child(1) > th:nth-child(1) > a:nth-child(1)').click(); err+=1
        time.sleep(0.5)
        prices = browser.find_elements_by_class_name("price"); err+=1
        ws['G61'] = int(prices[2].text.replace(',',"")); Message()
        ws['G62'] = int(prices[5].text.replace(',',"")); Message()
        ws['G63'] = int(prices[8].text.replace(',',"")); Message()

        Search('가루'); Message(); Message(); Message(); Message(); err+=1
        prices = browser.find_elements_by_class_name("price"); err+=1
        ws['G67'] = int(prices[2].text.replace(',',""))

        Search('정식'); err+=1
        browser.find_element_by_css_selector("#lostark-wrapper > div > main > div > div.deal > div.deal-contents > form > fieldset > div > div.detail > div.grade > div > div.lui-select__title").click()
        time.sleep(0.5); Message(); Message(); err+=1
        browser.find_element_by_css_selector("#lostark-wrapper > div > main > div > div.deal > div.deal-contents > form > fieldset > div > div.detail > div.grade > div > div.lui-select__option > label:nth-child(4)").click()
        time.sleep(0.5); Message(); Message(); err+=1
        browser.find_element_by_css_selector("#lostark-wrapper > div > main > div > div.deal > div.deal-contents > form > fieldset > div > div.bt > button.button.button--deal-submit").click()
        time.sleep(0.5); Message(); Message(); err+=1
        prices = browser.find_elements_by_class_name("price")
        names = browser.find_elements_by_class_name('name'); err+=1
        cnt_name = 0
        cnt=0
        j=69
        for i in prices:
            cnt+=1
            if cnt==3 :
                cnt=0
                cnt_name+=1
                try : 
                    tmp = browser.find_element_by_css_selector('#tbodyItemList > tr:nth-child({}) > td:nth-child(1) > div:nth-child(1) > span:nth-child(3)'.format(cnt_name)).text; err+=1
                    ws['F{}'.format(j)]=names[cnt_name].text; Message(); Message()
                    p=int(i.text.replace(',',""))
                    ws['G{}'.format(j)] = p
                    j+=1        
                except : None

        err = 9999
        t = time.strftime('%m%d_%H%M%S',time.localtime(time.time()))
        wb.save('{}.xlsm'.format(t)); Message()
        progress=99; Message()

        print('\n\n최저가 데이터 동기화 완료. \n{}.xlsm 파일에서 확인하세요.\n이 창은 5초 후 자동으로 꺼집니다.'.format(t))
        time.sleep(5)
        browser.quit()
        isError = False
        break
    except:
        print("\n최저가 탐색 중 오류 발생! ERROR : {}".format(err))
        print("동일한 에러가 반복될 경우 제작자에게 문의해주세요.")
        print("5초 후 처음부터 다시 탐색합니다.")
        time.sleep(5)

if isError:
    print("\n최저가 탐색 중 오류 발생! ERROR : {}".format(err))
    print("동일한 에러가 반복될 경우 제작자에게 문의해주세요.")
    print("에러가 5회 발생하여 5초 후 프로그램이 종료됩니다.")
    time.sleep(5)