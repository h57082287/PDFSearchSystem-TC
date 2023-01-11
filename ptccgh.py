import time
from PDFReader import PDFReader
from bs4 import BeautifulSoup
import os
import httpx
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox

# 澄清醫院
class PTCCGH():
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

        # 建立header
        self.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36'
        }
        # 建立payload
        self.payload = {
        'idType': '1',
        'patData': 'H125087083',
        'birthDate': '0870701',
        'Send': '',
        'csrf': 'e00716bb18df93a0e4d855493b2bb52877a06df1'
        }
        self.browser = browser

    # 2022/12/24加入
    def run(self):
        # 2022/12/24加入 (VPN 檢測)
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
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 澄清醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"澄清醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
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
        time.sleep(2)
        self.window.GUIRestart()
        self._endBrowser()
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        self.payload['patData'] = ID
        self.payload['birthDate'] = year + month + day
        while True:
            try:
                with httpx.Client(http2=True) as client :
                    respone = client.get('https://pt.ccgh.com.tw/register_search.htm')
                    soup = BeautifulSoup(respone.content,"html.parser")

                    # 讀取隱藏元素
                    self.payload['csrf'] = soup.find('input',{'name':'csrf'}).get('value')

                    # 發送請求
                    respone = client.post('https://pt.ccgh.com.tw/register_search_detail.htm',data=self.payload,headers=self.headers)
                    with open('reslut.html','w',encoding='utf-8') as f :
                        f.write(self._changeHTMLStyle(respone.text,"https://pt.ccgh.com.tw/"))
                break
            except httpx.ReadTimeout:
                try:
                    self.VPN.startVPN()
                    content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 澄清醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                    self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("取消掛號",(name + '_' + ID + '_澄清醫院.png')) :
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
    
    def _changeHTMLStyle(self,page_content,targer:str):
        soup = BeautifulSoup(page_content,"html.parser")
        # 替換屬性內容，強制複寫資源路徑(針對img)
        Tags = soup.find_all(['img'])
        for Tag in Tags:
            Tag['src'] = targer + Tag['src'] 
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            Tag['href'] = targer + Tag['href']
        return str(soup)

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")
