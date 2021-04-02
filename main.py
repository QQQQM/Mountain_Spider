#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   main.py    
@Contact :   qm_666@126.com

@Modify Time      @Author    @Version    @Desciption
------------      -------    --------    -----------
2021/3/30 19:57     QM          1.0         None
'''

# import lib
import time
import re
import xlsxwriter
from selenium import webdriver


class Mountain_Spider:
    def __init__(self):
        self.website = "https://www.env.go.jp/nature/satoyama/senteichi_ichiran.html"
        self.driver_path = "./chromedriver.exe"

    def main(self):
        xlsx_name = 'MountainSpider.xlsx'  # 以关键词名字命名 xlsx 表格
        workbook = xlsxwriter.Workbook(xlsx_name)  # 建立 xlsx 表格
        work_sheet = []
        work_sheet.append(workbook.add_worksheet('总表'))
        work_sheet[0].set_column(0, 13, 20)  # 设置宽度
        title_data = ['No.', '名称', 'ふりがな', '所在地', '選定基準１', '選定基準２', '選定基準３', '選定理由', '保全活用状況（取組状況）','活動主体', 'その他参考情報', '保全活用施策（実施状況等）' ]  # 设置标题文字
        work_sheet[0].write_row('A1', title_data)  # 写入title_data

        opt = webdriver.ChromeOptions()  # 选择为chrome浏览器
        opt.headless = True  # 选择为展现窗口模式
        driver = webdriver.Chrome(executable_path=self.driver_path, options=opt)  # 创建浏览器对象ss
        driver.maximize_window()  # 最大化窗口
        print("\n已成功创建浏览器对象！")
        driver.get(self.website)
        # time.sleep(2)
        print("已成功打开链接！")
        home_handle = driver.current_window_handle
        p_url = driver.find_elements_by_xpath("//li/a")
        url_box = []
        url_box.append("https://www.env.go.jp/nature/satoyama/01_hokkaido/hokkaido.html")
        for i in p_url:
            url_temp = i.get_attribute("href")
            if ('0' in url_temp or '1' in url_temp or '2' in url_temp or '3' in url_temp or '4' in url_temp) and '2015' not in url_temp:
                url_box.append(url_temp)
        print(url_box)
        driver.close()
        url_mother = []
        num1=0
        for url in url_box:
            url_mother.append(re.sub(r'[a-z]+\.html', "", url))
        for i in range(47):
            url_now = url_mother[i] + "no0" + str(i+1) + "-"
            print(url_now)
            driver = webdriver.Chrome(executable_path=self.driver_path, options=opt)  # 创建浏览器对象ss
            for j in range(50):
                j+=1
                num1+=1
                url_son = url_now + str(j) + ".html"
                print(url_son)
                driver.get(url_son)
                tbody_table = driver.find_elements_by_xpath("//tbody")
                if tbody_table == []:
                    break
                    driver.close()
                else:
                    tbody = tbody_table[0]
                    tr_table = tbody.find_elements_by_tag_name("tr")
                    colume_write = []
                    num2=0
                    for tr in tr_table:
                        content_table = re.split(r'\s', tr.text)
                        print(content_table)
                        if len(content_table) == 1:
                            content_table.append(" ")
                        colume_write.append(content_table[1])
                        work_sheet[0].write(num1, num2, content_table[1])
                        num2+=1
                    print(colume_write)
        workbook.close()
        driver.close()


if __name__ == '__main__':
    mountain_spider = Mountain_Spider()
    mountain_spider.main()




