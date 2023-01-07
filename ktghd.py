import time
from urllib.parse import quote
from bs4 import BeautifulSoup
import ddddocr
import os
from PDFReader import PDFReader
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import base64
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox


# 光田醫院：大甲院區
class KTGHD():
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
        # 各醫院新增項目
        self.idx = 0
        self.datalen = 0
        self.olddatalen = 0
        self.log = Log()
        if self.window.checkVal_AUVPNM.get() :
            self.VPN = VPN(self.window)
            VPNWindow(self.VPN)
            if not self.VPN.InstallationCkeck() :
                messagebox.showerror("VPN異常","請檢查您是否有安裝OpenVPN !!!")
                self.window.RunStatus = False
                self.browser.quit()
                os._exit(0)

    def run(self):
        for currentPage in range(self.currentPage-1,self.EndPage):
            if self._PDFData(currentPage) and self.window.RunStatus:
                for self.idx in range(self.currentNum-1,self.datalen) :
                    print(self.Data[self.idx])
                    print(self.window.RunStatus)
                    if ((currentPage != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                        content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 光田醫院：大甲院區\n當前第" + str(currentPage+1) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                        self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                        self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"光田醫院：大甲院區",self.Data[self.idx]['Born'],str(currentPage + 1),str(self.idx + 1))
                        time.sleep(2)
                    else:
                        break
            else:
                break
        try :
            self.VPN.stopVPN()
        except:
            pass
        self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
        time.sleep(2)
        self.window.GUIRestart()
        self._endBrowser()
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str) -> bool:
        self.browser.get("https://www.ktgh.com.tw/Reg_Cancel_index.asp?CatID=24&ModuleType=Y")
        selectOption = Select(self.browser.find_element(by=By.XPATH, value='//*[@id="Area"]'))
        selectOption.select_by_value("2")
        time.sleep(1)
        try:
            self.browser.find_element(by=By.XPATH, value='//*[@id="abgne_float_ad"]/span').click()
        except:
            pass
        time.sleep(1)
        self.browser.find_element(by=By.XPATH, value='//*[@id="Crtno"]').send_keys(ID)
        bornDate = month + day
        self.browser.find_element(by=By.XPATH, value='//*[@id="Brdte"]').send_keys(bornDate)
        Captcha = self._ParseCaptcha(self.browser.find_element(by=By.XPATH, value='//*[@id="imgnum"]'),self.browser,mode=1)
        time.sleep(5)
        self.browser.find_element(by=By.XPATH, value='//*[@id="ChkNum"]').send_keys(Captcha)
        time.sleep(5)
        self.browser.find_element(by=By.XPATH, value='//*[@id="reg"]/table/tbody/tr[6]/td/input[2]').click()
        time.sleep(3)
        
        while True:
            if '因系統安全驗證檢核，' in str(self.browser.page_source):
                self.window.setStatusText(content="因驗證碼錯誤，系統正重新查詢",x=0.2,y=0.7,size=20)
                self.browser.find_element(by=By.XPATH, value='//*[@id="Sizebox"]/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/input').click()
                self.browser.find_element(by=By.XPATH, value='//*[@id="reg"]/table/tbody/tr[4]/td[2]/table/tbody/tr[2]/td/a/b/font').click()
                self.browser.find_element(by=By.XPATH, value='//*[@id="ChkNum"]').clear()
                time.sleep(3)
                Captcha = self._ParseCaptcha(self.browser.find_element(by=By.XPATH, value='//*[@id="imgnum"]'),self.browser,mode=1)
                # print(Captcha)
                time.sleep(3)
                self.browser.find_element(by=By.XPATH, value='//*[@id="ChkNum"]').send_keys(Captcha)
                time.sleep(1)
                self.browser.find_element(by=By.XPATH, value='//*[@id="reg"]/table/tbody/tr[6]/td/input[2]').click()
                time.sleep(3)
            else:
                break

        with open('reslut.html','w', encoding='utf-8') as f :
            f.write(self._changeHTMLStyle(self.browser.page_source))

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("病情簡述",(name + '_' + ID + '_光田醫院：大甲院區.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','button','input','h1','h2','h3','h4','h5','span','b','font'])
        for tag in Tags :
            if tag.text == condition :
                found = True
                self.browser.save_screenshot(self.outputFile + '/' + fileName)
                break
        return found

    def _PDFData(self,currentPage) -> bool:
        # print("Current : " + str(self.currentPage) + "  End : " + str(self.EndPage))
        mPDFReader = PDFReader(self.window,self.filePath)
        status, self.Data,self.datalen = mPDFReader.GetData(currentPage)
        return status
    
    def _changeHTMLStyle(self,page_content):
        soup = BeautifulSoup(page_content,"html.parser")
        # 替換屬性內容，強制複寫資源路徑(針對img)
        Tags = soup.find_all(['img'])
        for Tag in Tags:
            Tag['src'] = 'https://www.ktgh.com.tw/' + Tag['src'] 
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            Tag['href'] = 'https://www.ktgh.com.tw/' + Tag['href']
        return str(soup)
    
    # 驗證碼辨識
    def _ParseCaptcha(self,element,driver,mode=1):
        if mode == 1 :
            image_base64 = driver.execute_script("\
                var ele = arguments[0];\
                var cnv = document.createElement('canvas');\
                cnv.width = ele.width;\
                cnv.height = ele.height;\
                cnv.fillStyle = '#FFFFFF';\
                cnv.getContext('2d').drawImage(ele,0,0);\
                return cnv.toDataURL('image/png').substring(22);\
            ",element)
        elif mode == 2 :
            image_base64 = driver.execute_script("\
                var ele = arguments[0];\
                var cnv = document.createElement('canvas');\
                cnv.width = 835;\
                cnv.height = 90;\
                cnv.getContext('2d').drawImage(ele,0,0);\
                return cnv.toDataURL('image/jepg').substring(22);\
            ",element)

        with open("captcha.png",'wb') as f :
            f.write(base64.b64decode(image_base64))
        
        with open("captcha.png",'rb') as f :
            img_bytes = f.read()
        orc = ddddocr.DdddOcr()
        result = orc.classification(img_bytes)
        return result
        # with open("VaildCode.png",'rb') as f :
        #     img_bytes = f.read()
        # orc = ddddocr.DdddOcr()
        # result = orc.classification(img_bytes)
        # os.remove("VaildCode.png")
        # return result

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")
