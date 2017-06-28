# coding = utf-8
import xlwt
from bs4 import BeautifulSoup
from selenium import webdriver
import os
import time

#浏览器启动
class SetUpBrowser(object):
    def start_browser(self,url):
        # 登录并通过浏览器进入待爬页面
        driver = webdriver.Firefox()
        driver.implicitly_wait(600)  # 设定最大等待时间
        driver.get(url)
        os.system('pause')  # 等待进入待爬取页面
        driver.switch_to_window(driver.window_handles[-1])  # 获取当前打开窗口句柄
        return driver  # 返回打开的浏览器对象

#爬取
class Spider(object):
    def craw(self,driver):
        # 生成excel以及构建字段
        row = 1  # 初始化写入元素的行
        book = xlwt.Workbook()  # 打开一个excel
        sheet1 = book.add_sheet('sheet1', cell_overwrite_ok=True)
        #构建表头
        heads = ['任务分类ID', '任务分类名称', '任务ID', '任务名称', '分发UID', '创建时间', '任务状态', '执行时间',
                 '完成时间', '责任人姓名', '责任人工号', '所属分组ID', '所属分组名称', '所属部门ID', '所属部门名称']  # 表头
        # 写入表头
        ii = 0
        for head in heads:
            sheet1.write(0, ii, head)
            ii += 1
        print(u'\n准备将数据存入表格...')

        # 记录当前页码，以翻页后页码不变作为循环结束条件
        page_current = driver.find_element_by_tag_name('strong').text

        # 开始爬取
        while True:
            # 解析为DOM树
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            # 获取表中元素
            content = soup.find_all('td',
                                    class_=['td-0', 'td-1', 'td-2', 'td-3', 'td-4', 'td-5', 'td-6', 'td-7', 'td-8',
                                            'td-9','td-10', 'td-11', 'td-12', 'td-13', 'td-14'])
            # 将信息放入一个list中,创建new_list(方便后续存入excel)(语法糖)
            data_list = []
            for data in content:
                data_list.append(data.text)
            new_list = [data_list[i:i + 15] for i in range(0, len(data_list), 15)]
            # 将表中元素写入excel表格
            for list in new_list:
                col = 0
                for data in list:
                    sheet1.write(row, col, data)
                    col += 1
                row += 1
            # 判断是否最后一页
            driver.find_element_by_class_name('next').click()
            page_next = driver.find_element_by_tag_name('strong').text
            if page_next != page_current:
                page_current = page_next
            else:
                break
        book.save('D:\\')
        print(u'\n录入成功！')

if __name__ == '__main__': # 让你写的脚本模块既可以导入到别的模块中用，另外该模块自己也可执行
    url = "https://login.alibaba-inc.com/verification/smsOrToken.htm?BACK_URL=http%3A%2F%2Fabc.aliyun-inc.co"
    "m%2F&CONTEXT_PATH=%2F&CLIENT_VERSION=0.3.9&APP_NAME=aliyun-business-center&CANCEL_CERT=true"
    browser = SetUpBrowser()
    driver = browser.start_browser(url)
    object_spider = Spider()
    object_spider.craw(driver)