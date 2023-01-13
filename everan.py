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
import random

# 太平長安醫院
class EVERAN():
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
        # if self.window.checkVal_AUVPNM.get() :
        #     self.VPN = VPN(self.window)
        #     VPNWindow(self.VPN)
        #     if not self.VPN.InstallationCkeck() :
        #         messagebox.showerror("VPN異常","請檢查您是否有安裝OpenVPN !!!")
        #         self.window.RunStatus = False
        #         self.browser.quit()
        #         os._exit(0)

        # payload2需要用到的時間
        self.loc_dt = datetime.datetime.today()
        self.new_dt = self.loc_dt + datetime.timedelta(days=365)
        # 建立header
        self.headers = {
                    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.47'
                }
        # 建立登入payload
        self.payload = {
                            'login': 'Y',
                            'cardid': 'H122222222',
                            'birthday': '1944-01-04',
                            'pattelphone': '',
                        }
        # 取得已掛號資訊的payload
        self.payload2 = {
                            'cardid': 'H12222222',
                            'birthday': '22222222',
                            'startdate': self.loc_dt.strftime("%Y%m%d"),
                            'enddate': self.new_dt.strftime("%Y%m%d")
                        }
        self.browser = browser

    def run(self):
        for self.page in range(self.currentPage-1,self.EndPage):
            if self.window.RunStatus:
                if self._PDFData(self.page):
                    for self.idx in range(self.currentNum-1,self.datalen) :
                        print(self.Data[self.idx])
                        if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 太平長安醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"太平長安醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            sec = random.randint(1, 5)
                            print(sec)
                            time.sleep(sec)
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
        self.payload['cardid'] = ID
        self.payload['birthday'] = year + '-' + month + '-' + day
        self.payload2['cardid'] = ID
        self.payload2['birthday'] = str(int(year) + 1911) + month + day
        with httpx.Client(http2=True) as client :
            # 進入網頁
            self.respone = client.post('http://www.everanhospital.com.tw:9099/OReg/GetPatInfo', data=self.payload, headers=self.headers, timeout=20)
            if(self.respone.json()[0]['msg'] == "病患不存在"):
                self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
                with open("reslut.html", "w", encoding="utf-8") as f:
                    f.write(self.respone.json()[0]['msg'])
            else:
                self.respone = client.get('http://www.everanhospital.com.tw:9099/OReg/ScheduledRecords', timeout=20)
                self.respone2 = client.post('http://www.everanhospital.com.tw:9099/OReg/GetEventsByCondition', data=self.payload2, headers=self.headers, timeout=20)
                self._JSONDataToHTML(self.respone2,self.respone.text)
        client.close()

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("預約成功",(name + '_' + ID + '_太平長安醫院.png')) :
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
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            Tag['href'] = "http://www.everanhospital.com.tw:9099/" + Tag['href']
        # 替換屬性內容，強制複寫資源路徑(針對img)
        Tags = soup.find_all(['img'])
        for Tag in Tags:
            Tag['src'] = "http://www.everanhospital.com.tw:9099/Content/onlineregister/logo.png"
        #增加掛號table
        Tags = soup.find(['table'])
        Tags['style'] = "width:100%;border:0px;font-size:2.0em;table-layout:fixed;"
        return str(soup)

    def _JSONDataToHTML(self,jsonData,page_content):
        page_content = self._changeHTMLStyle(page_content)
        soup = BeautifulSoup(page_content,"html.parser")
        datas = []
        datas = jsonData.json()
        for i in range(len(datas)):
            new_tr_tag = soup.new_tag("tr")
            new_tr_tag.attrs = {"class":"row"}
            target_tag = soup.tbody
            target_tag.append(new_tr_tag)

            target_tag = soup.find_all(['tr'])
            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = datas[i]['patname']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-3 col-md-2 col-lg-2"}
            new_td_tag.string = datas[i]['datetime']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-2 col-md-2 col-lg-2"}
            new_td_tag.string = datas[i]['sctename']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-2 col-md-2 col-lg-2"}
            new_td_tag.string = datas[i]['empname']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = str(datas[i]['visitno'])
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = datas[i]['regway']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = datas[i]['statuname']
            target_tag[i+1].append(new_td_tag)

        with open('reslut.html','w',encoding='utf-8') as f :
            f.write(str(soup))

    
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