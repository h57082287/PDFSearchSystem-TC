from bs4 import BeautifulSoup
import ddddocr
import os
import time
from PDFReader import PDFReader
from LogController import Log
# 2022/12/24加入
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert 


# 林新醫院
class LSHOSP():
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
        # 2022/12/24加入(各醫院新增項目)
        self.idx = 0
        self.page = 0
        self.datalen = 0
        self.log = Log()
        self.errorNum = 0
        self.maxError = 10
        self.url = "http://www.lshosp.com.tw:8001/OINetReg/OINetReg.Reg/Reg_RegConfirm1.aspx"
        self.browser = browser

    def run(self):
        # 2022/12/24加入 (VPN 檢測)
        if self.window.checkVal_AUVPNM.get() :
            self.VPN = VPN(self.window)
            print("a")
            VPNWindow(self.VPN)
            print("b")
            if not self.VPN.InstallationCkeck() :
                messagebox.showerror("VPN異常","請檢查您是否有安裝OpenVPN !!!")
                self.window.RunStatus = False
                self.browser.quit()
                os._exit(0)
            print("c")
        for self.page in range(self.currentPage-1,self.EndPage):
            if self.window.RunStatus:
                if self._PDFData(self.page):
                    for self.idx in range(self.currentNum-1,self.datalen) :
                        print(self.Data[self.idx])
                        if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 林新醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"林新醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            self.errorNum = 0
                            time.sleep(2)
                        else:
                            break
                self.currentNum = 1
            else:
                break
        try :
            self.VPN.stopVPN()
        except:
            pass
        self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
        self.errorNum = 0
        time.sleep(2)
        self.window.GUIRestart()
        self._endBrowser()
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        while True:
            try:
                self.browser.get(self.url)
                print(1)
                time.sleep(2)
                self.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtIDNOorPatientID"]').send_keys(ID)
                print(2)
                time.sleep(2)
                Select(self.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_dd1BirthM"]')).select_by_value(str(int(month)))
                print(3)
                time.sleep(2)
                Select(self.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_dd1BirthD"]')).select_by_value(str(int(day)))
                print(4)
                time.sleep(2)
                Captcha = self._ParseCaptcha4Img(self.browser.find_element(By.XPATH, '//*[@id="reserveTAB02-3"]/tbody/tr[2]/td[4]/img'))
                print(5)
                time.sleep(2)
                self.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_txtVerificationCode"]').send_keys(Captcha)
                print(6)
                time.sleep(2)
                self.browser.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnReg"]').click()
                print(7)
                time.sleep(2)
                if not self._CKCaptcha("Web", "驗證碼錯誤! 請輸入正確的驗證碼！"):
                    self.window.setStatusText(content="因驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                    time.sleep(2)
                    print("8-1")
                else:
                    print("8-2")
                    self.errorNum = 0
                    break
            except:
                print("發生錯誤即將重試(" + str(self.errorNum) + ")")
                self.window.setStatusText(content="~發生錯誤(" + str(self.errorNum) + ")，準備再次嘗試~\n~等候重新執行當前人員查詢~",x=0.3,y=0.8,size=12)
                if(self.errorNum >= self.maxError):
                    messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
                    os._exit(0)
                self.errorNum += 1
                time.sleep(5)

    def _startBrowser(self,name,ID):
        if self._Screenshot("就診序號",(name + '_' + ID + '_林新醫院.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','button','input','th','h1','h2','h3','h4','h5'])
        for tag in Tags :
            if tag.text == condition :
                found = True
                self.browser.save_screenshot(self.outputFile + '/' + fileName)
                break
        return found

    # 2022/12/14 加入
    def _PDFData(self,currentPage) -> bool:
        # print("Current : " + str(self.currentPage) + "  End : " + str(self.EndPage))
        mPDFReader = PDFReader(self.window,self.filePath)
        status, self.Data,self.datalen = mPDFReader.GetData(currentPage)
        print(status)
        return status
    
    # 檢查驗證是否成功
    def _CKCaptcha(self, CK_Method, msg):
        try:
            if CK_Method == "Alert":
                print("這是彈窗訊息 : " + Alert(self.browser).text.strip())
                print("這是條件訊息 : " + msg)
                if Alert(self.browser).text.strip() == msg:
                    Alert(self.browser).accept()
                    return False
                else:
                    print(Alert(self.browser).text)
                    Alert(self.browser).accept()
                    return True
            if CK_Method == "Web":
                html = str(BeautifulSoup(self.browser.page_source,"html.parser"))
                if msg in html:
                    return False
                return True
        except Exception as e :
            print("這是錯誤訊息 : " + str(e))
            return True
    
    # 驗證碼辨識
    def _ParseCaptcha4Img(self, element):
        image_base64 = self.browser.execute_script("\
                var ele = arguments[0];\
                var cnv = document.createElement('canvas');\
                cnv.width = ele.width;\
                cnv.height = ele.height;\
                cnv.fillStyle = '#FFFFFF';\
                cnv.getContext('2d').drawImage(ele,0,0);\
                return cnv.toDataURL('image/png').substring(22);\
            ",element)
        orc = ddddocr.DdddOcr()
        result = orc.classification(image_base64)
        return result

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")