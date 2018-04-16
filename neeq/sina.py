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
        numList = sheet1.col_values(2, 0)
        nameList = sheet1.col_values(1, 0)
        return dict(zip(numList, nameList))

    def readPdf(self, link, code, year, comDf):
        try:
            print('READING START')
            dfList = tabula.read_pdf(link, encoding='gbk', pages='all', guess = False, lattice = True, multiple_tables = True)
        except:
            return None
        
        dataTypeVis = [['资产总计', False], ['负债合计', False], ['基本每股收益', False]]
        dataType2Count = 0
        dataType3Count = 0
        strGudong = '股东名称(全称)'

        for df in dfList:
            
            for indexs in df.index:
                _dataName1 = ''
                _dataName2 = ''
                try:
                    dfLenth = len(df.loc[df.index[0]])
                    _dataName1 = str(df.loc[indexs].values[0]).strip().replace(' ','').replace('\r', '')
                    _dataName2 = str(df.loc[indexs].values[1]).strip().replace(' ','').replace('\r', '')
                except:
                    pass
                try:
                    if strGudong in _dataName1 or strGudong in _dataName2:
                        val = str(df.loc[indexs + 2].values[3]).strip().replace(' ','').replace('\r', '')
                        print('!!!!!!!!!!!!!Found it!!!!!!!!!!!!!!!!!!')
                        print(val)
                        if year == 2015:
                            comDf.loc[code, '第一股东持股比例 - 15'] = float(val) if val != '' and val != '-' and val != 'nan' else 0
                        elif year == 2016:
                            comDf.loc[code, '第一股东持股比例 - 16'] = float(val) if val != '' and val != '-' and val != 'nan' else 0
                        continue
                except:
                    pass
                if dataTypeVis[1][0] in _dataName1 or dataTypeVis[1][0] in _dataName2:
                    dataType2Count += 1
                if dataTypeVis[2][0] in _dataName1 or dataTypeVis[2][0] in _dataName2:
                    dataType3Count += 1
                for dType in dataTypeVis:
                    if dType[1] == False and (dType[0] in _dataName1 or dType[0] in _dataName2) and (dType[0] == '资产总计' or (dType[0] == '负债合计' and dataType2Count == 3) or (dType[0] == '基本每股收益' and dataType3Count == 3)):
                        dType[1] = True
                        try:
                            val1 = str(df.loc[indexs].values[dfLenth - 2]).strip()
                            val2 = str(df.loc[indexs].values[dfLenth - 1]).strip()
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
                
        print(comDf)
        return comDf

    def comProcess(self, que, num, numDict, comDataType):
        numInt = int(num)
        numStr = str(numInt)

        curURL = 'http://vip.stock.finance.sina.com.cn/corp/view/vCB_BulletinGather.php?stock_str=' + numStr + '&gg_date=&ftype=ndbg'

        print(curURL)
        driver = self.connect(curURL)

        driver.implicitly_wait(7)

        aXPATH = "//*[@id='wrap']/div[5]/table/tbody/tr/th/a"
        aList = driver.find_elements_by_xpath(aXPATH)

        pattern = re.compile(u'.*(2015|2016)年*年度报告(（已更正）|（修订版）|（更正修订）)*$')

        self.comDataFrame = pd.DataFrame(data = {'公司名称' : numDict[num]}, index = [numInt], columns = comDataType)

        for i in range(len(aList)):
            if pattern.match(aList[i].text) != None:
                link = aList[i + 1].get_attribute('href')
                print(aList[i].text)
                print(link)
                if '2015' in aList[i].text and link.endswith('PDF'):
                    self.comDataFrame = self.readPdf(link, num, 2015, self.comDataFrame)
                elif '2016' in aList[i].text and link.endswith('PDF'):
                    self.comDataFrame = self.readPdf(link, num, 2016, self.comDataFrame)

        driver.quit()

        que.put(self.comDataFrame)
        # print(self.comDataFrame)

    def getFiles(self):
        self.makedir('TEMP')
        xlsPath = os.path.join(self.basePath, 'testList.xlsx')
        numDict = self.readExcel(xlsPath)
        print(numDict)
        dataType = ['资产总计', '负债合计', '基本每股收益']

        comDataType = ['公司名称', '第一股东持股比例 - 15', '第一股东持股比例 - 16']
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
        exportXLSPath = os.path.join(self.basePath, 'DatasA.xlsx')
        totalDataFrame.to_excel(exportXLSPath, sheet_name = 'Data', index = True, header = True)
        return totalDataFrame

if __name__ == '__main__':
    obj = ExportDatas()
    obj.getFiles()
