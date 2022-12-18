import time
import requests
from PDFReader import PDFReader
from bs4 import BeautifulSoup
import os
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox

# 亞洲大學
class AUH():
    def __init__(self, browser, mainWindowsObj, S_Page:int=1, S_Num:int=1, E_Page:int=1, E_Num:int=5, outputFile:str=None, filePath:str=None) -> None:
        if E_Num == '':
            E_Num = 5
        self.code_rule = [400,500,503,404,408]
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
                        content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 亞洲大學附設醫院\n當前第" + str(currentPage) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                        self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                        self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"亞洲大學附設醫院",self.Data[self.idx]['Born'],str(currentPage + 1),str(self.idx + 1))
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

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        while True:
            try:
                respone = requests.get('https://appointment.auh.org.tw/cgi-bin/as/reg21.cgi?Tel=' + ID + '&sentbtn=%E7%A2%BA++++%E5%AE%9A&day=01&month=01&Year=088')
                if ("對不起!此ip查詢或取消資料次數過多;" in respone.text) or (respone.status_code in self.code_rule) :
                    raise requests.exceptions.ConnectTimeout("ip已被封鎖")
                break
            except requests.exceptions.ConnectTimeout:
                try:
                    self.VPN.startVPN()
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break

        with open("reslut.html",'wb') as f :
            f.write(respone.content)

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot(" 取消此筆掛號 ",(name + '_' + ID + '_亞洲大學.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','input','h1','h2','h3','h4','h5'])
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
    
    def _endBrowser(self):
        self.browser.quit()
        
    def __del__(self):
        print("物件刪除")
    
