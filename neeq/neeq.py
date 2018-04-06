from selenium import webdriver
from urllib.request import urlretrieve
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait  
from selenium.webdriver.support import expected_conditions as EC  
import os, time, xlrd, xlwt

class DownloadDatas():

    def __init__(self):
        self.url = 'http://www.neeq.com.cn/disclosure/announcement.html'
        self.basePath = os.path.dirname(__file__)

    def makedir(self, name):
        path = os.path.join(self.basePath, name)
        isExist = os.path.exists(path)
        if not isExist:
            os.makedirs(path)
            print('File has been created.')
        else:
            print('The file is existed.')
        # 切换到该目录下
        os.chdir(path)

    def connect(self, url):
        driver = webdriver.Chrome()
        driver.get(url)
        return driver

    def readExcel(self, path):
        workbook = xlrd.open_workbook(r'' + path)
        sheet1 = workbook.sheet_by_index(0)
        numList = sheet1.col_values(1, 1)
        nameList = sheet1.col_values(2, 1)
        return dict(zip(numList, nameList))

    def getFiles(self):
        driver = self.connect(self.url)
        self.makedir('Files')
        xlsPath = os.path.join(self.basePath, '创新层名单.xlsx')
        numDict = self.readExcel(r'C:\Users\Jack\Desktop\NEEQ\创新层名单.xlsx')

        driver.find_element_by_id('startDate').clear()
        driver.find_element_by_id('startDate').send_keys(u'2016-02-01')

        driver.find_element_by_id('endDate').clear()
        driver.find_element_by_id('endDate').send_keys(u'2017-05-01')

        driver.find_element_by_id('keyword').send_keys(u'年度报告')

        for num in numDict:
            num = int(num)
            driver.find_element_by_id('companyCode').clear()
            driver.find_element_by_id('companyCode').send_keys(str(num))
            driver.find_element_by_link_text('查询').click()

            try:
                tdList = WebDriverWait(driver, 10).until(EC.)
            aList = driver.find_elements_by_tag_name('td')
            for r in aList:
                try:
                    link = r.get_attribute('href')
                    if link.endswith('pdf'):
                        print(r.text)
                        print(link)
                        fileName = r.text + '.pdf'
                        urlretrieve(link, fileName)
                except:
                    pass


if __name__ == '__main__':
    obj = DownloadDatas()
    obj.getFiles()