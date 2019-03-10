#!/usr/bin/python3
# -*- coding: utf-8 -*-
import xlrd, xlwt, os
from xlutils.copy import copy
# from download import request
from bs4 import BeautifulSoup
import time
import platform
from selenium import webdriver

system = platform.system()
xlsPath = ''
#根据系统识别路径
if system == 'Darwin':#mac
    originPath = os.path.abspath('.')
    xlsPath = os.path.join(originPath,'loadData.xls')
elif system == 'Windows':
    originPath = 'C:/Users/Administrator/Desktop/Amzgetcomments'
    xlsPath = os.path.join(originPath,'loadData.xls')


#保存xls表格
def savexls():
    os.remove(xlsPath)
    workbookCopy.save(xlsPath)
#获取当前表格的信息
def get_sheet_mes():
    workbook = xlrd.open_workbook(xlsPath)
    workbookCopy = copy(workbook)

    sheet_name = workbook.sheet_names()[0]
    sheet_one = workbook.sheet_by_name(sheet_name)
    wordlist = sheet_one.col_values(1)
    selllist = sheet_one.col_values(0)
    return workbookCopy, wordlist, selllist

# 获取时间精确到秒s
def get_date():
    return time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))



# 初始化浏览器
browser = webdriver.Chrome()


def search(url,keyword):
    browser.get(url)
    inputs = browser.find_element_by_id('twotabsearchtextbox')# 找到输入框
    inputs.clear()#先清空搜索框

    inputs.send_keys(keyword)#传入关键词

    button = browser.find_element_by_xpath(".//*[@class='nav-input']")
    button.click()


def get_all_store(keyword):
    url = 'https://www.amazon.com/'
    search(url, keyword)
    #a-color-base s-line-clamp-4
    all_href = []
    all_stores = browser.find_elements_by_xpath(".//*[@class='a-color-base s-line-clamp-4']")
    for store in all_stores:
        hrefa = store.find_element_by_tag_name('a')
        href = hrefa.get_attribute('href')
        all_href.append(href)

    return all_href

def get_five_satrt(url):
    browser.get(url)

    store_name = browser.find_element_by_id('bylineInfo').get_attribute('textContent')

    try:
         # 五星标签 如果没有5星标签直接下一下
        five_satrt_a = browser.find_element_by_xpath(".//*[@class='a-size-base a-link-normal 5star histogram-review-count a-color-secondary']")

        five_start_url = five_satrt_a.get_attribute('href')
        five_satrt_a_value = five_satrt_a.get_attribute('textContent')
        five_satrt_num = five_satrt_a_value[:-1]
        #五星好评大于0个
        if int(five_satrt_num) > 0:
            browser.get(five_start_url)
            # a-section review aok-relative
            all_div = browser.find_elements_by_xpath(".//*[@class='a-section review aok-relative']")

            all_title = []
            all_time = []
            all_body = []
            for div in all_div:
                review_title = div.find_element_by_xpath(".//*[@class='cr-original-review-content']").get_attribute('textContent')
                review_time = div.find_element_by_xpath(".//*[@class='a-size-base a-color-secondary review-date']").get_attribute('textContent')
                review_body = div.find_element_by_xpath(".//*[@class='a-size-base review-text']").get_attribute('textContent')

                all_title.append(review_title)
                all_time.append(review_time)
                all_body.append(review_body)

            return True, store_name, all_title, all_time, all_body
        else:
            return False, store_name, None, None, None
    except Exception as e:
            return False, store_name, None, None, None
   






workbookCopy, wordlist, link_list = get_sheet_mes()
# sheet = workbookCopy.get_sheet(0)

all_href = get_all_store(link_list[1])
for href in all_href:
    ok, store_name, all_title, all_time, all_body = get_five_satrt(href)

    workbookCopy, wordlist, link_list = get_sheet_mes()
    sheet = workbookCopy.get_sheet(0)

    start_row = len(wordlist)
    if ok:
        for x in range(0,len(all_title)):
            print(x+start_row, x, start_row)
            sheet.write(start_row+x, 1, store_name)
            sheet.write(start_row+x, 2, all_time[x])
            sheet.write(start_row+x, 3, all_title[x])
            sheet.write(start_row+x, 4, all_body[x])
        #保存
        savexls()
    #暂停5s
    time.sleep(5)








