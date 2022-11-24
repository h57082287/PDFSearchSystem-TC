import time
import requests
from PDFReader import PDFReader
from bs4 import BeautifulSoup
import os

# 中國醫學大學豐原分院
class CMUH():
    def __init__(self, browser, mainWindowsObj, S_Page:int=1, S_Num:int=1, E_Page:int=1, E_Num:int=5, outputFile:str=None, filePath:str=None) -> None:
        if E_Num == '':
            E_Num = 5
        self.EndPage = int(E_Page)
        self.EndNum = int(E_Num)
        self.outputFile = outputFile
        self.window = mainWindowsObj
        self.filePath = filePath
        self.currentPage = int(S_Page)
        self.currentNum = int(S_Num)
        self.Data = []
        self.browser = browser

    def run(self):
        while True:
            if self._PDFData() and self.window.RunStatus:
                for persionData in self.Data :
                    print(persionData)
                    status = False
                    if (self.currentNum <= self.EndNum) and (self.currentPage <= self.EndPage) and self.window.RunStatus:
                        # 初診查詢
                        content = "姓名 : " + persionData['Name'] + "(初診)" + "\n身分證字號 : " + persionData['ID'] + "\n出生日期 : " + persionData['Born'] + "\n查詢醫院 : 中國醫學大學豐原分院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(persionData['Name'], persionData['ID'], "088","01","01")
                        status = self._startBrowser(persionData['Name'],persionData['ID'])
                        # 複診查詢
                        if not status:
                            content = "姓名 : " + persionData['Name']+ "(複診)"+ "\n身分證字號 : " + persionData['ID'] + "\n出生日期 : " + persionData['Born'] + "\n查詢醫院 : 中國醫學大學豐原分院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(persionData['Name'], persionData['ID'], persionData['Born'].split('/')[0],persionData['Born'].split('/')[1],persionData['Born'].split('/')[2])
                            self._startBrowser(persionData['Name'],persionData['ID'])
                        self.currentNum += 1
                        time.sleep(2)
                    else:
                        break
                self.currentNum = 1
                self.currentPage += 1
            else:
                self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
                self.window.GUIRestart()
                self._endBrowser()
                break
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        respone = requests.get('http://61.66.117.10/cgi-bin/eng/reg21.cgi?Tel=' + ID + '&sentbtn=%E7%A2%BA++++%E5%AE%9A&day=' + day + '&month=' + month + '&Year=' + year)
        with open('reslut.html','wb') as f :
            f.write(respone.content)

    def _startBrowser(self,name,ID) -> bool:
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        status = self._Screenshot(" 未看診 ",(name + '_' + ID + '_中國醫學大學豐原分院.png'))
        if status :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
        return status

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','input','h1','h2','h3','h4','h5','td'])
        for tag in Tags :
            # print(tag.text)
            if tag.text == condition :
                found = True
                self.browser.save_screenshot(self.outputFile + '/' + fileName)
                break
        return found

    def _PDFData(self) -> bool:
        # print("Current : " + str(self.currentPage) + "  End : " + str(self.EndPage))
        if (self.currentPage <= self.EndPage):
            mPDFReader = PDFReader(self.window,self.filePath)
            status, self.Data = mPDFReader.GetData(self.currentPage-1)
            return status
        else:
            return False
    
    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")