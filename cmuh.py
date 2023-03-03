import random
import time
import requests
from PDFReader import PDFReader
from bs4 import BeautifulSoup
import os
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
import selenium

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
        # 2022/12/24加入(各醫院新增項目)
        self.code_rule = [400,500,503,404,408]
        self.idx = 0
        self.page = 0
        self.datalen = 0
        self.log = Log()
        self.errorNum = 0
        self.maxError = 10

    def run(self):
        # 2022/12/25加入 (VPN 檢測)
        if self.window.checkVal_AUVPNM.get() :
            self.VPN = VPN(self.window)
            VPNWindow(self.VPN)
            if not self.VPN.InstallationCkeck() :
                messagebox.showerror("VPN異常","請檢查您是否有安裝OpenVPN !!!")
                self.window.RunStatus = False
                self.browser.quit()
                os._exit(0)
        for self.page in range(self.currentPage-1,self.EndPage):
            if self.window.RunStatus:
                if self._PDFData(self.page):
                    for self.idx in range(self.currentNum-1,self.datalen) :
                        print(self.Data[self.idx])
                        if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                            # 查詢狀態
                            Q_Status = False
                            # 初診查詢
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "(初診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 中國醫學大學豐原分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut_1(self.Data[self.idx]['Name'] + "(初診)", self.Data[self.idx]['ID'], "088","01","01")
                            Q_Status = self._startBrowser(self.Data[self.idx]['Name'] + "(初診)",self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'] + "(初診)","中國醫學大學豐原分院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            self._ClearCookie(self.browser)
                            time.sleep(2)
                            # 複診查詢
                            if not(Q_Status) and self.window.RunStatus:
                                content = "姓名 : " + self.Data[self.idx]['Name'] + "(複診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 中國醫學大學豐原分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                                self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                                self._getReslut_2(self.Data[self.idx]['Name'] + "(複診)", self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                                self._startBrowser(self.Data[self.idx]['Name'] + "(複診)",self.Data[self.idx]['ID'])
                                self.log.write(self.Data[self.idx]['Name'] + "(複診)",self.Data[self.idx]['ID'],"中國醫學大學豐原分院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                                self._ClearCookie(self.browser)
                                time.sleep(2)
                            else:
                                break
                        else:
                            break
                self.currentNum = 1
            else:
                break
        try :
            self.VPN.stopVPN()
            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 中國醫學大學豐原分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
        except:
            pass
        self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
        time.sleep(2)
        self.window.GUIRestart()
        self._endBrowser()
        del self

    def _getReslut_1(self,name:str, ID:str, year:str, month:str, day:str):
        while True:
            try:
                print("查詢網址 :" + 'http://61.66.117.10/cgi-bin/eng/reg21.cgi?Tel=' + ID + '&&sentbtn=%E7%A2%BA++++%E5%AE%9A&day=01&month=01&Year=88')
                self.browser.get('http://61.66.117.10/cgi-bin/eng/reg21.cgi?Tel=' + ID + '&&sentbtn=%E7%A2%BA++++%E5%AE%9A&day=01&month=01&Year=88')
                if ("對不起!此ip查詢或取消資料次數過多" in BeautifulSoup(self.browser.page_source, "html.parser").text.strip()) :
                    raise selenium.common.exceptions.TimeoutException("ip已被封鎖")
                time.sleep(random.randint(1,10))
                self.errorNum = 0
                break
            except selenium.common.exceptions.TimeoutException:
                try:
                    self.VPN.startVPN()
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break
            except:
                print("發生錯誤即將重試(" + str(self.errorNum) + ")")
                self._errorReTryTime()
                if(self.errorNum >= self.maxError):
                    messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
                    os._exit(0)
                self.errorNum += 1
                time.sleep(5)
    
    def _getReslut_2(self,name:str, ID:str, year:str, month:str, day:str):
        while True:
            try:
                print("查詢網址 : " + 'http://61.66.117.10/cgi-bin/eng/reg22.cgi?day=01&month=01&Year=088&CrtIdno='+ ID +'&sYear='+ year +'&sMonth='+ month +'&sDay='+ day +'&surebtn=%E7%A2%BA++%E5%AE%9A')
                self.browser.get('http://61.66.117.10/cgi-bin/eng/reg22.cgi?day=01&month=01&Year=088&CrtIdno='+ ID +'&sYear='+ year +'&sMonth='+ month +'&sDay='+ day +'&surebtn=%E7%A2%BA++%E5%AE%9A')
                print(BeautifulSoup(self.browser.page_source, "html.parser").text.strip())
                print("對不起!此ip查詢或取消資料次數過多" in BeautifulSoup(self.browser.page_source, "html.parser").text.strip())
                if ("對不起!此ip查詢或取消資料次數過多" in BeautifulSoup(self.browser.page_source, "html.parser").text.strip()) :
                    raise selenium.common.exceptions.TimeoutException("ip已被封鎖")
                time.sleep(random.randint(1,10))
                self.errorNum = 0
                break
            except selenium.common.exceptions.TimeoutException:
                try:
                    self.VPN.startVPN()
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break
            except:
                print("發生錯誤即將重試(" + str(self.errorNum) + ")")
                self._errorReTryTime()
                if(self.errorNum >= self.maxError):
                    messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
                    os._exit(0)
                self.errorNum += 1
                time.sleep(5)

    def _startBrowser(self,name,ID) -> bool:
        status = self._Screenshot("未看診",(name + '_' + ID + '_中國醫學大學豐原分院.png'))
        if status :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
        return status

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','input','h1','h2','h3','h4','h5','td','font'])
        for tag in Tags :
            print(tag.text.strip() + ':' + condition)
            if tag.text.strip() == condition :
                print("OKOK")
                found = True
                self.browser.save_screenshot(self.outputFile + '/' + fileName)
                break
        return found

    # 2022/12/14 加入
    def _PDFData(self,currentPage) -> bool:
        # print("Current : " + str(self.currentPage) + "  End : " + str(self.EndPage))
        mPDFReader = PDFReader(self.window,self.filePath)
        status, self.Data,self.datalen = mPDFReader.GetData(currentPage)
        return status
    
    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")
    
    def _errorReTryTime(self):
        self._ClearCookie(self.browser)
        self.browser.get("about:blank")
        min = random.randint(0,10)
        sec = 59
        for m in range(min, -1, -1):
            for s in range(sec, -1, -1):
                if m < 10:
                    mm = '0' + str(m)
                if s < 10:
                    ss = '0' + str(s)     
                self.window.setStatusText(content="~發生錯誤(" + str(self.errorNum) + ")，準備再次嘗試~\n~等候" + mm + ":" + ss + "重新執行~",x=0.3,y=0.8,size=12)
                time.sleep(1)
    
     # 清除快取
    def _ClearCookie(self,driver):
        driver.delete_all_cookies()