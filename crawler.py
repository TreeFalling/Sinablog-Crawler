# -*- coding:utf-8 -*-
from concurrent import futures  # 使用该模块实现进程池，用于编写异步多进程爬虫
from selenium import webdriver  # 模拟用户使用浏览器过程，动态抓取网页内容
from bs4 import BeautifulSoup  # 处理html信息
import xlwt  # 写入数据至Excel
import re


def wyh_experiment(date):
    # 创建.xls文件
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet(str(date) + '-' + str(date + 100), cell_overwrite_ok=True)

    # 解析date
    after_year = int(date / 10000)
    after_month = int((date - after_year * 10000) / 100)
    flag = 0
    if after_month == 12:
        before_year = after_year + 1
        before_month = 1
    else:
        before_year = after_year
        before_month = after_month + 1

    # 爬虫函数
    iterator = 0
    for page in range(1, 50):
        # 动态抓取并解析网页内容
        option = webdriver.ChromeOptions()
        # option.add_argument('--headless') # 调试时请将此语句变成注释，以便观察
        if after_month < 9:
            url = 'https://s.weibo.com/weibo?q=%E9%83%91%E5%B7%9E%E8%BD%BB%E5%B7%A5%E4%B8%9A%E5%A4%A7%E5%AD%A6' \
                  '&typeall=1&suball=1&timescope=' \
                  'custom:' + str(after_year) + '-0' + str(after_month) + '-01:' + str(before_year) + '-0' + str(
                before_month) + \
                  '-01' + '&Refer=g&page=' + str(page)
        elif after_month == 9:
            url = 'https://s.weibo.com/weibo?q=%E9%83%91%E5%B7%9E%E8%BD%BB%E5%B7%A5%E4%B8%9A%E5%A4%A7%E5%AD%A6' \
                  '&typeall=1&suball=1&timescope=' \
                  'custom:' + str(after_year) + '-0' + str(after_month) + '-01:' + str(before_year) + '-' + str(
                before_month) + \
                  '-01' + '&Refer=g&page=' + str(page)
        else:
            url = 'https://s.weibo.com/weibo?q=%E9%83%91%E5%B7%9E%E8%BD%BB%E5%B7%A5%E4%B8%9A%E5%A4%A7%E5%AD%A6' \
                  '&typeall=1&suball=1&timescope=' \
                  'custom:' + str(after_year) + '-' + str(after_month) + '-01:' + str(before_year) + '-' + str(
                before_month) + \
                  '-01' + '&Refer=g&page=' + str(page)
        driver = webdriver.Chrome(r"C:\chromedriver\chromedriver.exe", options=option)
        driver.get(url)
        web_data = driver.page_source
        soup = BeautifulSoup(web_data, 'lxml')

        # 若内容不足50页，则终止循环
        for x in soup.find_all('div', class_='card card-no-result s-pt20b40'):
            key = x.find_all('p')
            if key[0].string == '抱歉，未找到“郑州轻工业大学”相关结果。':
                break

        for k in soup.select("#pl_feedlist_index > div > div > div > div.card-feed > div.content > p"):
            # 使用正则表达式来规范爬取内容
            result = re.sub('\\<.*?\\>', '', str(k))
            result = re.sub('\s+', '', result).strip()
            if result[0] != '2':  # 取消日期信息的爬取（日期信息均以2019开头）
                iterator += 1
                sheet.write(iterator, 1, result.replace("", ""))

    driver.__exit__()  # 关闭当前界面
    workbook.save(str(date) + '-' + str(date + 100) + '.xls')  # 保存文件


# date数组里请放入间隔为一个月的日期，下面为实例
date = [20180501, 20180601, 20180701, 20180801, 20180901, 20181001, 20181101, 20181201, 20190201, 20190201, 20190301,
        20190401, 20190501, 20190601, 20190701, 20190801, 20190901, 20191001, 20191101, 20191201, 20200201, 20200201,
        20200301, 20200401
        ]
# 创建进程池并设置最大同时运行数
with futures.ThreadPoolExecutor(max_workers=1) as executor:
    for future in executor.map(wyh_experiment, date):
        print('Mission accomplished')  # 一个进程完成以后的提示
