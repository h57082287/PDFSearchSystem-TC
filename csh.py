import httpx
from bs4 import BeautifulSoup
import ddddocr
import os
import time
from PDFReader import PDFReader
import datetime
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox


# 中山醫
class CSH():
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

        # payload2需要用到的時間
        self.loc_dt = datetime.datetime.today()
        self.new_dt = self.loc_dt + datetime.timedelta(days=365)
        # 建立header
        self.headers = {    
                        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Mobile Safari/537.36',
                    }
        # 建立登入payload
        self.payload = {
                        'ToolkitScriptManager1' : 'ctl13|btnRegister',
                        'ToolkitScriptManager1_HiddenField' : '',
                        '__EVENTTARGET' : '',
                        '__EVENTARGUMENT' : '',
                        '__LASTFOCUS' : '',
                        '__VIEWSTATE' : '/wEPDwUINTgwNjg2NjgPZBYCAgMPZBYCAgkPZBYCZg9kFgYCGw8QZGQWAGQCHQ8QZGQWAWZkAh8PZBYCAgEPZBYCAgEPZBYCAgcPZBYCAgEPPCsADQBkGAIFCk11bHRpVmlldzEPD2RmZAUJR3JpZFZpZXcxD2dk5dWAueuvBsJUMBOozUzmtcX1Ug0=',
                        '__SCROLLPOSITIONX' : '0',
                        '__SCROLLPOSITIONY': '0',
                        '__PREVIOUSPAGE': 'Y-FBE3r6f27iDWHYfYWhX2I_SVTbXnloKkI072ZHH9RS-iP14oGUGXVKlDG37DVFdjdsznaLubXLYamj_PDvrBrjt_yfirkbg0z-zZkhrTatnXry0',
                        '__VIEWSTATEGENERATOR': 'B9F3ACFE',
                        '__EVENTVALIDATION': '/wEdAAYx6dgLzY/0SjpotPz1Y9B92eoiZVm6rgja/jcKL9Y4Ed2sXSaD1W3CCmvJLXVxJGMuJxBOmHVKfPCq/8YPb+k6tWGBV8+XxppQCkF63cYEoIlWBse8YG3rZW/ZgfSKpuyL3C4e2Z80bqIeZ6KKCmVg2vYteNaRNbq8667n9h2q9w==',
                        'rblZone': 'A',
                        'tbIdNo': 'H122222222',
                        'tbBirthday': '00000000',
                        'tbValid': '0000',
                        'hfRegMode': '',
                        'hfZone': 'A',
                        '__ASYNCPOST': 'true',
                        'btnRegister': '確定',
                    }
        self.browser = browser

    def run(self):
        for currentPage in range(self.currentPage-1,self.EndPage):
            if self._PDFData(currentPage) and self.window.RunStatus:
                for self.idx in range(self.currentNum-1,self.datalen) :
                    print(self.Data[self.idx])
                    print(self.window.RunStatus)
                    if ((currentPage != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                        content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 中山醫\n當前第" + str(currentPage+1) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                        self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                        self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"中山醫",self.Data[self.idx]['Born'],str(currentPage + 1),str(self.idx + 1))
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
        self.payload['tbIdNo'] = ID
        self.payload['tbBirthday'] = str(int(year) + 1911) + month + day

        with httpx.Client(http2=True) as client :
            self.respone = client.get('https://sysint.csh.org.tw/Register/SearchReg.aspx')
            # print(respone.status_code)
            soup = BeautifulSoup(self.respone.content,"html.parser")
            time.sleep(1)

            # 獲取隱藏元素
            self.payload['ToolkitScriptManager1_HiddenField'] = soup.find('form',{'id':'form1'}).find('input',{'name':'ToolkitScriptManager1_HiddenField'}).get('value')
            self.payload['__VIEWSTATE'] = soup.find('input',{'id':'__VIEWSTATE'}).get('value')
            self.payload['__PREVIOUSPAGE'] = soup.find('input',{'id':'__PREVIOUSPAGE'}).get('value')
            self.payload['__VIEWSTATEGENERATOR'] = soup.find('input',{'id':'__VIEWSTATEGENERATOR'}).get('value')
            self.payload['__EVENTVALIDATION'] = soup.find('input',{'id':'__EVENTVALIDATION'}).get('value')

            while True:
                # 請求驗證碼
                self.respone = client.get('https://sysint.csh.org.tw/Register/ValidateCookie.aspx')
                with open('VaildCode.png','wb') as f :
                    f.write(self.respone.content)
                self.payload['tbValid'] = self._ParseCaptcha()

                # 發送登入請求
                self.respone = client.post('https://sysint.csh.org.tw/Register/SearchReg.aspx', headers=self.headers, data=self.payload)
                soup = BeautifulSoup(self.respone.content, "html.parser")
                if('對不起，您輸入的驗證碼有誤，請再輸入一次，謝謝!' not in soup):
                    break
                
            self._changeHTMLStyle(self.respone.content)

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("取消此筆掛號",(name + '_' + ID + '_中山醫.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','button','input','h1','h2','h3','h4','h5','td'])
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
            Tag['src'] = 'https://sysint.csh.org.tw' + Tag['src'] 
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            Tag['href'] = 'https://sysint.csh.org.tw' + Tag['href']

        table = soup.find_all(['table'])
        with open('reslut.html','w',encoding='utf-8') as f :
            f.write(str(table[2]))
    
    def _CKCaptcha(self,page_content,contentType,keyWord) -> bool:
        soup = BeautifulSoup(page_content,"html.parser")
        if keyWord in str(soup.find_all(contentType)):
            return True
        return False

    # 驗證碼辨識
    def _ParseCaptcha(self):
        with open("VaildCode.png",'rb') as f :
            img_bytes = f.read()
        orc = ddddocr.DdddOcr()
        result = orc.classification(img_bytes)
        os.remove("VaildCode.png")
        return result

    # 各醫院新增項目
    def _ChangingIPCK(self):
        while(self.ChangeIPNow):
            pass
        self.ChangeIPNow = False

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")