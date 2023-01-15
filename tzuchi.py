import time
from PDFReader import PDFReader
from bs4 import BeautifulSoup
import os
import httpx
import ddddocr
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox

# L122782985
# 710208
        
# 慈濟醫院台中分院
class TZUCHI():
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
        self.maxError = 10
        self.errorNum = 0

        # 建立header
        self.header = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"
        }

        self.Q_Payload = {
            "__LASTFOCUS": "",
            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",
            "__VIEWSTATE": "/wEPDwUKLTgwMDQwOTQwNQ8WBB4Ec0xvYwIHHgdzSG9zcE5vBQJUQxYCZg9kFgICAw9kFgQCAQ9kFgQCAQ9kFgQCBw9kFgICAQ8PFgIeC05hdmlnYXRlVXJsBShodHRwczovL3RhaWNodW5nLnR6dWNoaS5jb20udHcvc2NoZWR1bGUvZGQCCQ9kFgICAQ8PFgIeB1Zpc2libGVoZGQCAg9kFgJmD2QWAgIBDw8WAh4ISW1hZ2VVcmwFHWltYWdlcy/oh7rkuK3mhYjmv5/phqvpmaIuanBnZGQCAw9kFgJmD2QWAmYPZBYCAgEPZBYMAgMPDxYCHgRUZXh0BZ4DMS7niZnnp5Et5pif5pyf5YWt5p2O5piO5YSS6Yar5bir54K66b2S6aGO55+v5q2j56eR54m557SE6ZaA6Ki644CCPHA+Mi7oi6XmgqjpnIDopoHlgrflj6Pmj5vol6XvvIzkuJTnhKHlpJbnp5HploDoqLrkuYvmmYLmrrXvvIzoq4vnm7TmjqXoh7PmgKXoqLrlrqTmjpvomZ/kuKboqqrmmI7opoHlgrflj6Pmj5vol6XvvIzmgKXoqLrln7fooYzlgrflj6Pmj5vol6XmlLbosrvoiIfkuIDoiKzploDoqLrmlLbosrvnm7jlkIzjgII8L3A+PHA+My7oq4vlsLHoqLrnl4XmgqPlpoLmlrwxMOaXpeWFp+abvuiHs0g3TjnmtYHmhJ/nl4XkvovnmbznlJ/lnLDljYDml4XpgYrvvIzkuJTmnInnmbznh5Llj4rlkrPll73nrYnnl4fni4DvvIzmlrzlsLHphqvmmYLljbPkuLvli5XlkYrnn6Xnm7jpl5zml4XpgYrlj7Llj4rnl4fni4DjgII8L3A+ZGQCBg9kFgICAQ9kFgICBQ88KwALAGQCCA8PFgIfBQUw5L2b5pWZ5oWI5r+f6Yar55mC6LKh5ZyY5rOV5Lq66Ie65Lit5oWI5r+f6Yar6ZmiZGQCCg8PFgIfBQUy5Zyw5Z2AOuWPsOS4reW4gua9reWtkOWNgOixkOiIiOi3r+S4gOautTY244CBODjomZ9kZAIMDw8WAh8FBRXpm7voqbHvvJowNC0zNjA2LTA2NjZkZAIODw8WAh8FBSnmhI/opovlj43mmKDkv6HnrrHvvJp0Y21haWxAdHp1Y2hpLmNvbS50d2RkGAEFJGN0bDAwJENvbnRlbnRQbGFjZUhvbGRlcjEkTXVsdGlWaWV3MQ8PZGZkMH8iA2oPJkRUTIe4gcVvGADknqWOU0OLxBLXswPeVak=",
            "__VIEWSTATEGENERATOR": "9265AEB6",
            "__EVENTVALIDATION": "/wEdAAf9/9QMpJphC9PVmnOSSxkQIcVBFhJvDOedlDgOyFOsVnqNruCzOJOM3QgMNFviaBIypIRMWfsPhbBSwpFEmuVbxzMMSP6wPpKcct5IXVrK3Pm5vKEn+J0ZI7Cx0WC73tzkPI7bL93seMnBn71LTUOw8hOYTxjseIynGk5ZbRSQKFpsmmP7vuvMEHuqNO9ihpI=",
            "D2": "#",
            "ctl00$ContentPlaceHolder1$txtMRNo": "L122782985",
            "ctl00$ContentPlaceHolder1$txtVCode": "04130",
            "ctl00$ContentPlaceHolder1$btnQry": "查詢",
        }

        self.V_Payload = {
            "__LASTFOCUS": "",
            "__EVENTTARGET": "",
            "__EVENTARGUMENT":"", 
            "__VIEWSTATE": "/wEPDwUKLTgwMDQwOTQwNQ8WBB4Ec0xvYwIHHgdzSG9zcE5vBQJUQxYCZg9kFgICAw9kFgQCAQ9kFgQCAQ9kFgQCBw9kFgICAQ8PFgIeC05hdmlnYXRlVXJsBShodHRwczovL3RhaWNodW5nLnR6dWNoaS5jb20udHcvc2NoZWR1bGUvZGQCCQ9kFgICAQ8PFgIeB1Zpc2libGVoZGQCAg9kFgJmD2QWAgIBDw8WAh4ISW1hZ2VVcmwFHWltYWdlcy/oh7rkuK3mhYjmv5/phqvpmaIuanBnZGQCAw9kFgJmD2QWAmYPZBYCAgEPZBYMAgMPDxYCHgRUZXh0BZ4DMS7niZnnp5Et5pif5pyf5YWt5p2O5piO5YSS6Yar5bir54K66b2S6aGO55+v5q2j56eR54m557SE6ZaA6Ki644CCPHA+Mi7oi6XmgqjpnIDopoHlgrflj6Pmj5vol6XvvIzkuJTnhKHlpJbnp5HploDoqLrkuYvmmYLmrrXvvIzoq4vnm7TmjqXoh7PmgKXoqLrlrqTmjpvomZ/kuKboqqrmmI7opoHlgrflj6Pmj5vol6XvvIzmgKXoqLrln7fooYzlgrflj6Pmj5vol6XmlLbosrvoiIfkuIDoiKzploDoqLrmlLbosrvnm7jlkIzjgII8L3A+PHA+My7oq4vlsLHoqLrnl4XmgqPlpoLmlrwxMOaXpeWFp+abvuiHs0g3TjnmtYHmhJ/nl4XkvovnmbznlJ/lnLDljYDml4XpgYrvvIzkuJTmnInnmbznh5Llj4rlkrPll73nrYnnl4fni4DvvIzmlrzlsLHphqvmmYLljbPkuLvli5XlkYrnn6Xnm7jpl5zml4XpgYrlj7Llj4rnl4fni4DjgII8L3A+ZGQCBQ9kFgRmD2QWBAIBDw8WAh8DaGQWBAIDDw8WAh8FBQpMMTIyNzgyOTg1ZGQCCQ8PFgIfBQUFNzY1OTBkZAIDDw8WAh8DZ2RkAgEPZBYCAgUPPCsACwBkAgcPDxYCHwUFMOS9m+aVmeaFiOa/n+mGq+eZguiyoeWcmOazleS6uuiHuuS4reaFiOa/n+mGq+mZomRkAgkPDxYCHwUFMuWcsOWdgDrlj7DkuK3luILmva3lrZDljYDosZDoiIjot6/kuIDmrrU2NuOAgTg46JmfZGQCCw8PFgIfBQUV6Zu76Kmx77yaMDQtMzYwNi0wNjY2ZGQCDQ8PFgIfBQUp5oSP6KaL5Y+N5pig5L+h566x77yadGNtYWlsQHR6dWNoaS5jb20udHdkZBgBBSRjdGwwMCRDb250ZW50UGxhY2VIb2xkZXIxJE11bHRpVmlldzEPD2RmZLbU2uGALotnaNDi4hj5THW6so4JrBipHNGQFqOlMlxD",
            "__VIEWSTATEGENERATOR": "9265AEB6",
            "__EVENTVALIDATION": "/wEdAAcOaUHjJ7jeFMrbRAnJjNBHIcVBFhJvDOedlDgOyFOsVnqNruCzOJOM3QgMNFviaBIypIRMWfsPhbBSwpFEmuVbvcJx1V0mBDwP/M1+X07GqoraaQaFnXKaLuO1sKRm4e/kLgzbJ9fFW/8j4U/WT5la+NLr2pIQzchxs+Ghp0DudvVOSm5lld/LYwa6dl3z0z8=",
            "D2": "#",
            "ctl00$ContentPlaceHolder1$txtBirthday": "0710208",
            "ctl00$ContentPlaceHolder1$Button2": "第二道驗證碼確認",
        }
        self.browser = browser

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
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"慈濟醫院台中分院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            time.sleep(2)
                        else:
                            break
                self.currentNum = 1
            else:
                break
        try :
            self.VPN.stopVPN()
            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
        except:
            pass
        self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
        time.sleep(2)
        self.window.GUIRestart()
        self._endBrowser()
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        self.Q_Payload['ctl00$ContentPlaceHolder1$txtMRNo'] = ID
        self.V_Payload['ctl00$ContentPlaceHolder1$txtBirthday'] = year + month + day
        while True:
            try:
                with httpx.Client(http2=True) as client :
                    # 請求查詢頁
                    respone = client.get("https://app.tzuchi.com.tw/tchw/opdreg/RegQryCancel.aspx?Loc=TC")
                    soup = BeautifulSoup(respone.content,"html.parser")

                    # 獲取隱藏資料
                    self.Q_Payload["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
                    self.Q_Payload["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
                    self.Q_Payload["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")

                    # 請求驗證碼
                    while True :
                        respone = client.get("https://app.tzuchi.com.tw/tchw/opdreg/ImageCode.aspx")
                        with open("VaildCode.png","wb") as f :
                            f.write(respone.content)
                        self.Q_Payload["ctl00$ContentPlaceHolder1$txtVCode"] = self._ParseCaptcha()

                        # 發送查詢請求
                        respone = client.post("https://app.tzuchi.com.tw/tchw/opdreg/RegQryCancel.aspx?Loc=TC",data=self.Q_Payload,headers=self.header)
                        soup = BeautifulSoup(respone.content,"html.parser")
                        if self._CKCaptcha(respone.content,"span","驗證失敗!"): 
                            self.window.setStatusText(content="驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                            content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            time.sleep(1)
                        elif not self._SencordCK(str(soup),"請輸入第二道驗證碼【個人出生日期】！"):
                            with open("reslut.html","w",encoding="utf-8") as f :
                                f.write(self._changeHTMLStyle(respone.content,"https://app.tzuchi.com.tw/tchw/opdreg/",""))
                            break
                        else :
                            # 獲取隱藏資料
                            self.V_Payload["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
                            self.V_Payload["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
                            self.V_Payload["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
                            # 發送第二次驗證請求
                            respone = client.post("https://app.tzuchi.com.tw/tchw/opdreg/RegQryCancel.aspx?Loc=TC",data=self.V_Payload,headers=self.header)
                            with open("reslut.html","w",encoding="utf-8") as f :
                                f.write(self._changeHTMLStyle(respone.content,"https://app.tzuchi.com.tw/tchw/opdreg/",""))
                            break
                break
            except httpx.ConnectTimeout:
                print("發生時間例外")
                self.window.setStatusText(content="~連線超時，啟動VPN~",x=0.3,y=0.75,size=14)
                self.errorNum = 0
                try:
                    self.VPN.startVPN()
                    content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                    self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break
            except AttributeError:
                self.window.setStatusText(content="~網頁請求回應不完全，即將重試(" + str(self.errorNum) + ")~",x=0.3,y=0.75,size=14)
                self.errorNum += 1
                if(self.errorNum > self.maxError):
                    self.errorNum = 0
                    try:
                        self.VPN.startVPN()
                        content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                    except:
                        messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                        self.window.Runstatus = False
                        break
                print("AttributeError")
                time.sleep(5)
            except httpx.ReadTimeout:
                print("ReadTimeout")
                self.window.setStatusText(content="~網頁讀取超時，啟動VPN~",x=0.3,y=0.75,size=14)
                self.errorNum = 0
                try:
                    self.VPN.startVPN()
                    content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                    self.window.setStatusText(content=content,x=0.3,y=0.75,size=14)
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break
            except httpx.ReadError:
                print("ConnectTimeout")
                self.window.setStatusText(content="~網頁讀取錯誤，啟動VPN~",x=0.3,y=0.75,size=14)
                self.errorNum = 0
                try:
                    self.VPN.startVPN()
                    content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 慈濟醫院台中分院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                    self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("看診號碼",(name + '_' + ID + '_慈濟醫院台中分院.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)

    def _Screenshot(self,condition:str,fileName:str) -> bool:
        found = False
        soup = BeautifulSoup(self.browser.page_source,"html.parser")
        Tags = soup.find_all(['a','button','input','th','td','h1','h2','h3','h4','h5'])
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
    
    def _changeHTMLStyle(self,page_content,targer1:str,targer2:str):
        soup = BeautifulSoup(page_content,"html.parser")
        # 替換屬性內容，強制複寫資源路徑(針對img)
        Tags = soup.find_all(['img'])
        for Tag in Tags:
            try:
                Tag['src'] = targer1 + Tag['src'] 
            except:
                pass
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            try:
                Tag['href'] =  targer1 + Tag['href']
            except:
                pass
        # 替換屬性內容，強制複寫資源路徑(針對input)
        Tags = soup.find_all(['input'])
        for Tag in Tags:
            try:
                Tag['src'] = targer1 + Tag['src']
            except:
                pass
        # 替換屬性內容，強制複寫資源路徑(針對input)
        Tags = soup.find_all(['tr'])
        for Tag in Tags:
            try:
                Tag['class'] = ""
            except:
                pass
        return str(soup)
    
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
    
    def _SencordCK(self,page_content,keyword) -> bool:
        if keyword in page_content :
            return True
        return False

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")