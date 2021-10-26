from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import openpyxl
import re
import pyperclip

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

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
browser = webdriver.Chrome('chromedriver.exe', options=options)
browser.get('https://lostark.game.onstove.com/Market')

wb = openpyxl.load_workbook('base.xlsm', read_only=False, keep_vba=True)
ws = wb['거래소 최저가']

login_type = wb['검색']['I5'].internal_value
uid = wb['검색']['I6'].internal_value
upw = wb['검색']['I7'].internal_value


Login()

print('\n최저가 데이터 받아오는 중...')






browser.find_element_by_css_selector('.main-category > li:nth-child(8) > a:nth-child(1)').click()
browser.find_element_by_css_selector('.is-active > ul:nth-child(2) > li:nth-child(1) > a:nth-child(1)').click()
time.sleep(1.5)
browser.find_element_by_css_selector('#itemList > thead:nth-child(3) > tr:nth-child(1) > th:nth-child(1) > a:nth-child(1)').click()

cnt=0
j=1
for k in range(1,5) :
    if k>1 :
        browser.execute_script('paging.page({});'.format(k))
    time.sleep(0.7)
    prices = browser.find_elements_by_class_name("price")
    names = browser.find_elements_by_class_name('name')
    cnt_name = 1
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
    j+=1

browser.find_element_by_css_selector('.main-category > li:nth-child(6) > a:nth-child(1)').click()
browser.find_element_by_css_selector('.is-active > ul:nth-child(2) > li:nth-child(1) > a:nth-child(1)').click()
time.sleep(1)
browser.find_element_by_css_selector('#itemList > thead:nth-child(3) > tr:nth-child(1) > th:nth-child(1) > a:nth-child(1)').click()

cnt=0
j=1
for k in range(1,7) :
    if k>1 :
        browser.execute_script('paging.page({});'.format(k))
    time.sleep(0.7)
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
    j+=1

Search('하 융')
browser.find_element_by_css_selector('#itemList > thead:nth-child(3) > tr:nth-child(1) > th:nth-child(1) > a:nth-child(1)').click()
time.sleep(0.5)
prices = browser.find_elements_by_class_name("price")
ws['G61'] = int(prices[2].text.replace(',',""))
ws['G62'] = int(prices[5].text.replace(',',""))
ws['G63'] = int(prices[8].text.replace(',',""))

Search('가루')
prices = browser.find_elements_by_class_name("price")
ws['G67'] = int(prices[2].text.replace(',',""))

Search('정식')
browser.find_element_by_css_selector("#lostark-wrapper > div > main > div > div.deal > div.deal-contents > form > fieldset > div > div.detail > div.grade > div > div.lui-select__title").click()
time.sleep(0.5)
browser.find_element_by_css_selector("#lostark-wrapper > div > main > div > div.deal > div.deal-contents > form > fieldset > div > div.detail > div.grade > div > div.lui-select__option > label:nth-child(4)").click()
time.sleep(0.5)
browser.find_element_by_css_selector("#lostark-wrapper > div > main > div > div.deal > div.deal-contents > form > fieldset > div > div.bt > button.button.button--deal-submit").click()
time.sleep(0.5)
prices = browser.find_elements_by_class_name("price")
names = browser.find_elements_by_class_name('name')
cnt_name = 0
cnt=0
j=69
for i in prices:
    cnt+=1
    if cnt==3 :
        cnt=0
        cnt_name+=1
        try : 
            tmp = browser.find_element_by_css_selector('#tbodyItemList > tr:nth-child({}) > td:nth-child(1) > div:nth-child(1) > span:nth-child(3)'.format(cnt_name)).text
            ws['F{}'.format(j)]=names[cnt_name].text
            p=int(i.text.replace(',',""))
            ws['G{}'.format(j)] = p
            j+=1        
        except : None


t = time.strftime('%m%d_%H%M%S',time.localtime(time.time()))
wb.save('{}.xlsm'.format(t))

print('최저가 데이터 동기화 완료. {}.xlsm 파일에서 확인하세요.\n이 창은 5초 후 자동으로 꺼집니다.'.format(t))
time.sleep(5)
browser.quit()