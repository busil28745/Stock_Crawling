
import time
from random import randint

import pandas as pd
import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

MAX_SLEEP_TIME = 5

url = "https://seibro.or.kr/websquare/control.jsp?w2xPath=/IPORTAL/user/bond/BIP_CNTS03005V.xml&menuNo=88"
excel_file = "CB___2016-2021.xlsx"

df_excel = pd.read_excel(excel_file, sheet_name = 'CB 행사내역', usecols= [2])

fncNm = []
#엑셀에서 종목명만 리스트에 저장
for idx in df_excel.index:
    fncNm.append(df_excel.loc[idx, '종목명' ])
#print(fncNm)

#크롤링 시작
browser = webdriver.Chrome('C://chromedriver/chromedriver.exe')
wait = WebDriverWait( browser, 10 )

browser.get(url)
for fncItm in fncNm:
    innerList = []
    # 검색할 종목
    trg_name = fncItm
    # 검색어 입력
    wait.until( EC.element_to_be_clickable( (By.ID, 'INPUT_SN2') ) )
    browser.find_element_by_xpath('//*[@id="INPUT_SN2"]').send_keys( trg_name + Keys.ENTER )
    # 팝업으로 이동
    wait.until( EC.element_to_be_clickable( (By.ID, 'iframeIsin') ) )
    iframe = browser.find_element_by_xpath('//*[@id="iframeIsin"]')
    browser.switch_to.frame(iframe)
    # 목표 종목 클릭
    wait.until( EC.element_to_be_clickable( (By.ID, 'isinList_0_ISIN_ROW')))
    browser.find_element_by_xpath('//*[@id="isinList_0_ISIN_ROW"]').click() # 검색 목록 중 첫번째 행 선택
    # 조회 클릭
    browser.switch_to.default_content()
    wait.until( EC.element_to_be_clickable( (By.ID, 'group186' ) ) )
    browser.find_element_by_xpath('//*[@id="group186"]').click()
    # 내용물 로드 기다림
    while(True):
        time.sleep(.1)
        content_text = browser.find_element_by_xpath('//*[@id="txt1_ISSU_DT"]').text
        innerList.append(content_text)
        if len( content_text ) > 0:
            print( content_text ) # debug
            break
    # data parsing
    '''
    parsing code
    test
    '''

    # refresh
    rand_value = randint(2, MAX_SLEEP_TIME) # 2초~MAX
    time.sleep(rand_value) # 서버 부하 방지
    browser.refresh()

browser.quit()