import time
from urllib.parse import quote
import httpx
from bs4 import BeautifulSoup
import ddddocr
import os
from PDFReader import PDFReader
import json
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
import selenium
import random

# 大里仁愛醫院
class JJAH():
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
        self.url = ""
        self.browser = browser

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
                            self.Val[2]['value'] = 'Y'
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "(初診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            Q_Status = self._getReslut(self.Data[self.idx]['Name'] + "(初診)", self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'] + "(初診)",self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'] + "(初診)",self.Data[self.idx]['ID'],"大里仁愛醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            time.sleep(2)
                            # 複診查詢
                            if not(Q_Status) and self.window.RunStatus:
                                self.Val[2]['value'] = 'N'
                                content = "姓名 : " + self.Data[self.idx]['Name'] + "(複診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                                self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                                Q_Status = self._getReslut(self.Data[self.idx]['Name'] + "(複診)", self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                                self._startBrowser(self.Data[self.idx]['Name'] + "(複診)",self.Data[self.idx]['ID'])
                                self.log.write(self.Data[self.idx]['Name'] + "(複診)",self.Data[self.idx]['ID'],"大里仁愛醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
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
        except:
            pass
        self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
        time.sleep(2)
        self.window.GUIRestart()
        self._endBrowser()
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str) -> bool:
        
        return status

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("取消掛號",(name + '_' + ID + '_大里仁愛醫院.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','button','input','h1','h2','h3','h4','h5'])
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
        return status
    
    # def _changeHTMLStyle(self,page_content,targer:str):
    #     soup = BeautifulSoup(page_content,"html.parser")
    #     # 替換屬性內容，強制複寫資源路徑(針對img)
    #     Tags = soup.find_all(['img'])
    #     for Tag in Tags:
    #         Tag['src'] = targer + Tag['src'] 
    #     # 替換屬性內容，強制複寫資源路徑(針對link)
    #     Tags = soup.find_all(['link'])
    #     for Tag in Tags:
    #         Tag['href'] = targer + Tag['href']
    #     return str(soup)
    
    def _JSONDataToHTML(self,jsonData) -> bool:
        tableData = ''
        datas = []
        if jsonData['codeError'] != 'Y':
            if jsonData['QueryList'] != '':
                datas = json.loads(jsonData['QueryList'])
            for data in datas:
                tableData += '<tr><td>大里仁愛醫院</td>\
                                <td>' + data['opdDate'][0:4] + '-' + data['opdDate'][4:6] + '-' + data['opdDate'][6:8] + '(' + data['week'] + ')</td>\
                                <td></td>\
                                <td>' + data['deptName'] + '</td>\
                                <td>' + data['doctorName'] + '</td>\
                                <td>' + data['opdTimeID'] + '</td>\
                                <td></td>\
                                <td></td>\
                                <td>' + data['estiTime'] + '</td>\
                                <td><a class="btn cancel openCheckCancelReg" href="#modal-cancel" data-lity="">取消掛號</a></td>\
                                </tr>'
            html = '<html>\
                        <head>\
                            <link href="https://www.jah.org.tw/JCHReg/Styles/style.css" rel="stylesheet">\
                            <link href="https://www.jah.org.tw/JCHReg/Styles/sweetalert.css" rel="stylesheet">\
                        </head>\
                        <span><sapn id="name">初診病患</sapn>  君，您的掛號資料：</span>\
                        <table class="table wide10 two-tone query">\
                            <thead>\
                                <tr>\
                                    <th>院區</th>\
                                    <th>看診日期</th>\
                                    <th>診別</th>\
                                    <th>科別</th>\
                                    <th>醫師姓名	</th>\
                                    <th>序號</th>\
                                    <th>門診視訊</th>\
                                    <th>就診地點</th>\
                                    <th>預定看診時間</th>\
                                    <th>是否取消掛號</th>\
                                </tr>\
                            </thead>\
                            <tbody id="regViewList">' + tableData + '\
                            </tbody>\
                        </table>\
                    </html>'
            with open('reslut.html','w',encoding='utf-8') as f :
                f.write(html)
            return True
        else:
            return False
    
    # 驗證碼辨識
    def _ParseCaptcha(self):
        with open("VaildCode.png",'rb') as f :
            img_bytes = f.read()
        orc = ddddocr.DdddOcr()
        result = orc.classification(img_bytes)
        os.remove("VaildCode.png")
        return result

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")
