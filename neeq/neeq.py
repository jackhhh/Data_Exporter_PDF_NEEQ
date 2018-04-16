from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait  
import os, time, xlrd, xlwt, re, tabula
import pandas as pd
import numpy as np
import multiprocessing
import openpyxl

class ExportDatas():

    def __init__(self):
        self.url = 'http://www.neeq.com.cn/disclosure/announcement.html'
        self.basePath = os.path.dirname(__file__)

    def makedir(self, name):
        path = os.path.join(self.basePath, name)
        isExist = os.path.exists(path)
        if not isExist:
            os.makedirs(path)
            print('TEMP has been created.')
        else:
            print('TEMP is existed.')
        os.chdir(path)

    def connect(self, url):
        option = webdriver.ChromeOptions()
        option.add_argument('headless')
        driver = webdriver.Chrome(chrome_options = option)
        driver.get(url)
        return driver

    def readExcel(self, path):
        workbook = xlrd.open_workbook(r'' + path)
        sheet1 = workbook.sheet_by_index(0)
        numList = sheet1.col_values(1, 1)
        nameList = sheet1.col_values(2, 1)
        return dict(zip(numList, nameList))

    def readPdf(self, link, code, year, comDf):
        dfList = tabula.read_pdf(link, encoding='gbk', pages='all', guess = False, lattice = True,  multiple_tables = True)

        dataTypeVis = [['可供出售金融资产', False], ['持有至到期投资', False], ['长期股权投资', False], ['投资性房地产', False], ['资产总计', False], ['营业收入', False], ['公允价值变动收益', False], ['投资收益', False], ['汇兑收益', False], ['三、营业利润', False], ['五、净利润', False], ['基本每股收益', False], ['销售商品、提供劳务收到的现金', False]]
        for df in dfList:
            for indexs in df.index:
                _dataName1 = ''
                _dataName2 = ''
                try:
                    _dataName1 = str(df.loc[indexs].values[0]).strip().replace(' ','').replace('\r', '')
                    _dataName2 = str(df.loc[indexs].values[1]).strip().replace(' ','').replace('\r', '')
                    dfLenth = len(df.loc[df.index[0]])
                except:
                    pass
                for dType in dataTypeVis:
                    if dataTypeVis[0][1] == False and (dType[0] == '基本每股收益' or dType[0] == '营业收入' or dType[0] == '资产总计' or dType[0] == '长期股权投资'):
                        continue
                    if dType[1] == False and ((dType[0] in _dataName1) or (dType[0] in _dataName2)):
                        dType[1] = True
                        try:
                            val1 = str(df.loc[indexs].values[dfLenth - 2]).strip().replace(' ','').replace('\r', '')
                            val2 = str(df.loc[indexs].values[dfLenth - 1]).strip().replace(' ','').replace('\r', '')
                        except:
                            pass
                        try:
                            if year == 2015:
                                _columnA = str(dType[0] + ' - 15末')
                                _columnB = str(dType[0] + ' - 15初')
                                comDf.loc[code, _columnA] = float(val1.replace(',', '')) if val1 != '' and val1 != '-' and val1 != 'nan' else 0
                                comDf.loc[code, _columnB] = float(val2.replace(',', '')) if val2 != '' and val2 != '-' and val2 != 'nan' else 0
                            elif year == 2016:
                                _columnA = str(dType[0] + ' - 16末')
                                _columnB = str(dType[0] + ' - 16初')
                                comDf.loc[code, _columnA] = float(val1.replace(',', '')) if val1 != '' and val1 != '-' and val1 != 'nan' else 0
                                comDf.loc[code, _columnB] = float(val2.replace(',', '')) if val2 != '' and val2 != '-' and val2 != 'nan' else 0
                        except:
                            pass

        return comDf

    def comProcess(self, que, num, numDict, comDataType):
        numInt = int(num)
        numStr = str(numInt)

        driver = self.connect(self.url)
        driver.find_element_by_id('startDate').clear()
        driver.find_element_by_id('startDate').send_keys(u'2016-02-01')
        driver.find_element_by_id('endDate').clear()
        driver.find_element_by_id('endDate').send_keys(u'2017-05-01')
        driver.find_element_by_id('keyword').send_keys(u'年度报告')
        driver.find_element_by_id('companyCode').clear()
        driver.find_element_by_id('companyCode').send_keys(numStr)
        driver.find_element_by_link_text('查询').click()

        driver.implicitly_wait(7)
        
        numsXPATH = '%s%s%s' % ("//*[@id='companyTable']/tr[1]/td[1]/font[text()='", numStr, "']")
        numList = driver.find_elements_by_xpath(numsXPATH)

        aXPATH = "//*[@id='companyTable']/tr/td/a"
        aList = driver.find_elements_by_xpath(aXPATH)
        
        pattern = re.compile(u'.*(2015|2016)年*年度报告(（已更正）)*$')

        comDataFrame = pd.DataFrame(data = {'公司名称' : numDict[num]}, index = [numInt], columns = comDataType)

        for r in aList:
            try:
                link = r.get_attribute('href')
                if link.endswith('pdf') and pattern.match(r.text) != None:
                    print(r.text)
                    print(link)
                    if '2015' in r.text:
                        comDataFrame = self.readPdf(link, num, 2015, comDataFrame)
                    elif '2016' in r.text:
                        comDataFrame = self.readPdf(link, num, 2016, comDataFrame)
                    
            except:
                pass

        driver.quit()

        que.put(comDataFrame)
        print(comDataFrame)

    def getFiles(self):
        self.makedir('TEMP')
        xlsPath = os.path.join(self.basePath, 'CtestList.xlsx')
        numDict = self.readExcel(xlsPath) 

        dataType = ['可供出售金融资产', '持有至到期投资', '长期股权投资', '投资性房地产', '资产总计', '营业收入', '公允价值变动收益', '投资收益', '汇兑收益', '三、营业利润', '五、净利润', '基本每股收益', '销售商品、提供劳务收到的现金']

        comDataType = ['公司名称']
        for _dataType in dataType:
            comDataType.append(_dataType + ' - 15初')
        for _dataType in dataType:
            comDataType.append(_dataType + ' - 15末')
        for _dataType in dataType:
            comDataType.append(_dataType + ' - 16初')
        for _dataType in dataType:
            comDataType.append(_dataType + ' - 16末')

        totalDataFrame = pd.DataFrame(columns = comDataType)

        manager = multiprocessing.Manager()
        que = manager.Queue()
        pool = multiprocessing.Pool(4)
        
        for num in numDict:
            try:
                pool.apply_async(self.comProcess, (que, num, numDict, comDataType, ))
            except:
                pass

        pool.close()
        pool.join()

        while not que.empty():
            try:
                totalDataFrame = totalDataFrame.append(que.get())
            except:
                pass

        print(totalDataFrame)
        exportXLSPath = os.path.join(self.basePath, 'DatasC.xlsx')
        totalDataFrame.to_excel(exportXLSPath, sheet_name = 'Data', index = True, header = True)
        return totalDataFrame

if __name__ == '__main__':
    obj = ExportDatas()
    obj.getFiles()
