import httpx
import tkinter
from bs4 import BeautifulSoup
import ddddocr
import os
import time

from selenium.common.exceptions import TimeoutException
from PDFReader import PDFReader
import datetime
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
import random
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import selenium


# 豐原醫院
class FYH():
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
        self.errorNum = 0
        self.maxError = 10
        self.Data = []
        self.browser = browser
        # 各醫院新增項目
        self.idx = 0
        self.datalen = 0
        self.olddatalen = 0
        self.log = Log()
        self.reTry = 0
        self.maxTry = 5

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

    def run(self):
        
        # if self.window.checkVal_AUVPNM.get() :
        #     self.VPN = VPN(self.window)
        #     print("a")
        #     VPNWindow(self.VPN)
        #     print("b")
        #     if not self.VPN.InstallationCkeck() :
        #         messagebox.showerror("VPN異常","請檢查您是否有安裝OpenVPN !!!")
        #         self.window.RunStatus = False
        #         self.browser.quit()
        #         os._exit(0)
        #     print("c")
                
        for self.page in range(self.currentPage-1,self.EndPage):
            if self.window.RunStatus:
                if self._PDFData(self.page):
                    for self.idx in range(self.currentNum-1,self.datalen) :
                        print(self.Data[self.idx])
                        if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 豐原醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"豐原醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            self._ClearCookie(self.browser)
                            sec = random.randint(1, 5)
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
        # self._endBrowser()
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        self.payload['cardid'] = ID
        self.payload['birthday'] = year + '-' + month + '-' + day
        self.payload2['cardid'] = ID
        self.payload2['birthday'] = str(int(year) + 1911) + month + day

        try:
            print("A")
            # delay = random.randint(1, 5)
            with httpx.Client(http2=True) as client :
                print("B")
                time.sleep(8)
                # 進入網頁
                self.respone = client.post('https://nreg.fyh.mohw.gov.tw/OReg/GetPatInfo', data=self.payload, headers=self.headers, timeout=20)
                print("C")
                if(self.respone.json()[0]['msg'] == "病患不存在"):
                    print("D")
                    self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
                    with open("reslut_豐原醫院.html", "w", encoding="utf-8") as f:
                        f.write(self.respone.json()[0]['msg'])
                    print("E")
                else:
                    print("F")
                    time.sleep(8)
                    self.respone = client.get('https://nreg.fyh.mohw.gov.tw/OReg/ScheduledRecords', timeout=20)
                    print("G")
                    time.sleep(8)
                    self.respone2 = client.post('https://nreg.fyh.mohw.gov.tw/OReg/GetEventsByCondition', data=self.payload2, headers=self.headers, timeout=20)
                    print("H")
                    self._JSONDataToHTML(self.respone2,self.respone.text)
                    print("I")
            print("J")
            client.close()
            time.sleep(15)
        except TimeoutException as error:
                print("偵測到瀏覽器超時異常，系統即將啟動VPN並重試")
                print(error)
                time.sleep(5)
                # try:
                #     self.VPN.startVPN()
                # except:
                #     tkinter.messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                #     self.window.Runstatus = False
                #     break
        except Exception as e :
            print("這是錯誤訊息 : " + str(e))
            print("發生錯誤即將重試(" + str(self.errorNum) + ")")
            self._errorReTryTime()
            if(self.errorNum >= self.maxError):
                tkinter.messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
                os._exit(0)
            self.errorNum += 1
            time.sleep(5)
        
        # self.browser.get("https://nreg.fyh.mohw.gov.tw/OReg/HomePage#")
        # time.sleep(3)
        # self.browser.find_element(by=By.XPATH, value='//*[@id="btn-login"]').click()
        # time.sleep(1)
        # self.browser.find_element(by=By.XPATH, value='//*[@id="user-cardid"]').send_keys(ID)
        # birthday = year + month + day
        # self.browser.find_element(by=By.XPATH, value='//*[@id="user-birthday"]').send_keys(birthday)
        # time.sleep(5)
        # self.browser.find_element(by=By.XPATH, value='//*[@id="login-confirm"]').click()
        # time.sleep(1)
        # status  = False
        # result1 = EC.alert_is_present()(self.browser)  #檢查是否為初診
        # print(result1)
        # if result1:
        #     alert = self.browser.switch_to_alert()
        #     alert.accept()
        #     self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
        # else:
        #     try:
        #         if '此證號已有資料病歷，請輸入正確。或請與本院聯絡，謝謝！' in str(self.browser.page_source):
        #             self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
        #             status = True
        #     except:
        #         pass
        #     if not status :
        #         time.sleep(5)
        #         self.browser.find_element(by=By.XPATH, value='//*[@id="scheduledrecords"]').click()
        #         time.sleep(5)
        #         if self._Screenshot("預約成功",(name + '_' + ID + '_豐原醫院.png')) :
        #             self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        #         else:
        #             self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)  
        #         self.browser.find_element(by=By.XPATH, value='//*[@id="btn-logout"]').click()
        # self.browser.get("https://nreg.fyh.mohw.gov.tw/OReg/HomePage#")

    def _startBrowser(self,name,ID):
        print(os.path.dirname(os.path.abspath(__file__)))
        print((os.path.dirname(os.path.abspath(__file__))).replace('\_internal',''))
        print(str(os.path.dirname(os.path.abspath(__file__))).replace('\_internal',''))
        self.browser.get(r'file:///' + (os.path.dirname(os.path.abspath(__file__))).replace('\_internal','') + '/reslut_豐原醫院.html')
        if self._Screenshot("預約成功",(name + '_' + ID + '_豐原醫院.png')) :
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
        mPDFReader = PDFReader(self.window,self.filePath)
        status, self.Data,self.datalen = mPDFReader.GetData(currentPage)
        return status
    
    def _changeHTMLStyle(self,page_content):
        soup = BeautifulSoup(page_content,"html.parser")
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            Tag['href'] = "https://nreg.fyh.mohw.gov.tw/" + Tag['href']
        # 替換屬性內容，強制複寫資源路徑(針對img)
        Tags = soup.find_all(['img'])
        for Tag in Tags:
            Tag['src'] = "https://nreg.fyh.mohw.gov.tw/Content/onlineregister/logo.png"
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
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-2 col-md-2 col-lg-2"}
            new_td_tag.string = datas[i]['visitroom']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = str(datas[i]['visitno'])
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = datas[i]['regarrivaltime']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = datas[i]['regway']
            target_tag[i+1].append(new_td_tag)

            new_td_tag = soup.new_tag("td")
            new_td_tag.attrs = {"style":"text-align:center", "class":"col-xs-1 col-md-1 col-lg-1"}
            new_td_tag.string = datas[i]['statuname']
            target_tag[i+1].append(new_td_tag)

        with open('reslut_豐原醫院.html','w',encoding='utf-8') as f :
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
    
    def _errorReTryTime(self):
        self._ClearCookie(self.browser)
        try:
            self.browser.get("about:blank")
        except:
            pass
        min = 5
        sec = 30
        for m in range(min, -1, -1):
            for s in range(sec, -1, -1):
                ss = str(s)
                mm = str(m)
                if m < 10:
                    mm = '0' + str(m)
                if s < 10:
                    ss = '0' + str(s)     
                self.window.setStatusText(content="~發生錯誤(" + str(self.errorNum) + ")，準備再次嘗試~\n~等候" + mm + ":" + ss + "重新執行~",x=0.3,y=0.8,size=12)
                time.sleep(1)
                
    # 清除快取
    def _ClearCookie(self,driver):
        try:
            driver.delete_all_cookies()
        except:
            pass

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")