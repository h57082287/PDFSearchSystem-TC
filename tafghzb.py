import time
from urllib.parse import quote
import httpx
from bs4 import BeautifulSoup
import ddddocr
import os
from PDFReader import PDFReader
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
import random
import socket
import base64
from selenium.webdriver.common.by import By

# 國軍醫院-中清
class TAFGHZB():
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
        self.ErrorNum = 0
        self.ErrorMax = 10
        # 各醫院新增項目
        self.idx = 0
        self.datalen = 0
        self.olddatalen = 0
        self.log = Log()

        # 建立header
        self.headers = {    
                'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Mobile Safari/537.36',
            }
        # 建立payload
        self.payload = {
                'idtype' : 'id1',
                'idno' : 'H122222222',
                'year' : '',
                'month' : '',
                'day' : '',
                'birth' : '888888',
                'vcode' : '',
                '__RequestVerificationToken': '',
            }
        self.browser = browser

    def run(self):
        socket.setdefaulttimeout(20)
        for self.page in range(self.currentPage-1,self.EndPage):
            if self.window.RunStatus:
                if self._PDFData(self.page):
                    for self.idx in range(self.currentNum-1,self.datalen) :
                        print(self.Data[self.idx])
                        if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 國軍醫院-中清\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            if self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2]):
                                # self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                                self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"國軍醫院-中清",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            sec = random.randint(1, 5)
                        else:
                            break
                        self.ErrorNum = 0
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
        status = False
        self.payload['idno'] = ID
        self.payload['birth'] = str(int(year)) + month + day
        with httpx.Client(http2=True) as client :
            # 利用迴圈自動重試
            while True:
                try:
                    self.browser.get("https://web-reg-server.803.org.tw/816/webreg/book_query")
                    time.sleep(3)
                    self.browser.find_element(by=By.XPATH, value='//*[@id="idno"]').send_keys(ID)
                    time.sleep(2)
                    self.browser.find_element(by=By.XPATH, value='//*[@id="birthType"]').click()
                    time.sleep(2)
                    bornText = str(int(year)) + month + day
                    self.browser.find_element(by=By.XPATH, value='//*[@id="birth"]').send_keys(bornText)
                    print(bornText)
                    Captcha = self._ParseCaptcha(element=self.browser.find_element(by=By.XPATH, value='//*[@id="codeimg"]'),driver=self.browser,mode=2)
                    time.sleep(2)
                    print(Captcha)
                    self.browser.find_element(by=By.XPATH, value='//*[@id="vcode"]').send_keys(Captcha)
                    time.sleep(2)
                    self.browser.find_element(by=By.XPATH, value='//*[@id="Send"]').click()
                    time.sleep(5)

                    while True:
                        # 檢查驗證碼是否輸入正確
                        checkCaptcha = BeautifulSoup(self.browser.page_source,"html.parser")
                        # 將element.Tag轉型成字串
                        buffer = []
                        for buff in checkCaptcha.find_all('div',{'class':'modal-body'}):
                            buffer.append(buff.text.replace(' ','').replace('\n',''))
                        # 檢查錯誤用
                        for b in buffer:
                            print(b)
                        if ('驗證碼輸入錯誤！' in buffer) :
                            self.window.setStatusText(content="因驗證碼錯誤，系統正重新查詢",x=0.2,y=0.7,size=20)
                            self.browser.find_element(by=By.XPATH, value='//*[@id="alertMe"]/div/div/div[1]/button/span').click()
                            time.sleep(5)
                            Captcha = self._ParseCaptcha(element=self.browser.find_element(by=By.XPATH, value='//*[@id="codeimg"]'),driver=self.browser,mode=2)
                            time.sleep(2)
                            print(Captcha)
                            self.browser.find_element(by=By.XPATH, value='//*[@id="vcode"]').send_keys(Captcha)
                            time.sleep(2)
                            self.browser.find_element(by=By.XPATH, value='//*[@id="Send"]').click()
                            time.sleep(2)
                        else:
                            time.sleep(2)
                            break
                    time.sleep(5)

                    if self._Screenshot("我要取消",(name + '_' + ID + '_國軍醫院-中清.png')) :
                        self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
                        self.browser.find_element(by=By.XPATH, value='/html/body/main/div/div/div/div/div[3]/div/div/div/div[2]/a').click()
                    else:
                        try:
                            self.browser.find_element(by=By.XPATH, value='/html/body/main/div/div/div/div/div[3]/div/div/div/div[2]/a').click()
                        except:
                            pass
                        self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)                            

                    # print("f1")
                    # # 獲取登入網頁回應
                    # time.sleep(5)
                    # self.respone = client.get('https://web-reg-server.803.org.tw/816/WebReg/book_query', timeout=20)
                    # soup = BeautifulSoup(self.respone.content,"html.parser")

                    # print("f2")
                    # # 獲取隱藏元素
                    # self.payload['__RequestVerificationToken'] = soup.find('form',{'method':'post'}).find('input',{'name':'__RequestVerificationToken'}).get('value')
                    
                    # print("f3")
                    # # 請求驗證碼
                    # time.sleep(5)
                    # self.respone = client.get('https://web-reg-server.803.org.tw/816/captcha-img', timeout=20)
                    # with open('VaildCode.png','wb') as f :
                    #     f.write(self.respone.content)
                    # self.payload['vcode'] = self._ParseCaptcha()

                    # print("f4")
                    # # 發送登入請求
                    # self.respone = client.post('https://web-reg-server.803.org.tw/816/WebReg/book_query', headers=self.headers, data=self.payload, timeout=20)
                    # time.sleep(10)
                    # # 檢查是否登入成功(有登入成功此網站會回應302)
                    # if(self.respone.status_code == 302):
                    #     print("f5")
                    #     # 查詢掛號紀錄
                    #     time.sleep(5)
                    #     self.respone = client.get('https://web-reg-server.803.org.tw/816/WebReg/book_detail')
                    #     with open('reslut.html','w', encoding='utf-8') as f :
                    #         f.write(self._changeHTMLStyle(self.respone.content))
                    #     status = True
                    #     break
                    # else:
                    #     print("f6")
                    #     # 沒有登入成功的話先檢查有沒有病歷資料
                    #     soup = BeautifulSoup(self.respone.content,"html.parser")
                    #     if("無符合病歷資料，請填寫初診資料以建立初次掛號" in str(soup)):
                    #         self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
                    #         with open("reslut.html", "w", encoding="utf-8") as f:
                    #             f.write("病患不存在")
                    #         status = True
                    #         break
                    #     # 有病歷資料的話即為驗證碼輸入錯誤，進行重試
                    #     print("f7")
                    #     self.window.setStatusText(content="驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                    #     self.ErrorNum += 1
                    #     print("驗證碼重試次數 : " + str(self.ErrorNum))
                    #     if(self.ErrorNum >= self.ErrorMax):
                    #         break
                    #     time.sleep(5)
                    #     content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 國軍醫院-中清\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx +1) + "筆"
                    #     self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                except Exception as e:
                    print(e)
                    print("發生錯誤即將重試(" + str(self.ErrorNum) + ")")
                    self._errorReTryTime()
                    if(self.ErrorNum >= self.ErrorMax):
                        messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
                        os._exit(0)
                    self.ErrorNum += 1
                    time.sleep(5)
                time.sleep(15)
                return status
        # return status

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','button','input','h1','h2','h3','h4','h5','span'])
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
            Tag['src'] = 'https://web-reg-server.803.org.tw' + Tag['src'] 
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            Tag['href'] = 'https://web-reg-server.803.org.tw' + Tag['href']
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

    # 各醫院新增項目
    def _ChangingIPCK(self):
        while(self.ChangeIPNow):
            pass
        self.ChangeIPNow = False

    def _endBrowser(self):
        self.browser.quit()

    def _errorReTryTime(self):
        self._ClearCookie(self.browser)
        self.browser.get("about:blank")
        min = random.randint(0,10)
        sec = 59
        for m in range(min, -1, -1):
            for s in range(sec, -1, -1):
                ss = str(s)
                mm = str(m)
                if m < 10:
                    mm = '0' + str(m)
                if s < 10:
                    ss = '0' + str(s)     
                self.window.setStatusText(content="~發生錯誤(" + str(self.ErrorNum) + ")，準備再次嘗試~\n~等候" + mm + ":" + ss + "重新執行~",x=0.3,y=0.8,size=12)
                time.sleep(1)
    
     # 清除快取
    def _ClearCookie(self,driver):
        try:
            driver.delete_all_cookies()
        except:
            pass

    def __del__(self):
        print("物件刪除")
