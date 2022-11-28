import httpx
from bs4 import BeautifulSoup
import ddddocr
import os
import time
from PDFReader import PDFReader

# 部立台中醫院
class MOHW():
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
        # 建立header
        self.header = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"
        }

        self.payloadM = {
            "ctl00$ScriptManager1":"ctl00$ContentPlaceHolder1$UpdatePanel2|ctl00$ContentPlaceHolder1$dd1BirthM",
            "__EVENTTARGET":"ctl00$ContentPlaceHolder1$dd1BirthM",
            "__EVENTARGUMENT": "",
            "__LASTFOCUS":"", 
            "__VIEWSTATE": "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",
            "__VIEWSTATEGENERATOR": "C8D299FD",
            "__VIEWSTATEENCRYPTED": "",
            "__EVENTVALIDATION":"ayDMtuYQwCBau/+sb75iWxAxDu47YDkEHWwI+582V/XcuNOrIWfWY3uBkxCOTkgbAHI3pAKJyp7QapB/166GSV39wczVypzGg4yshF+Cdz65g41V3XCkRXWTNUTWOTOADFUo1f3tTnquhXwrOrkvsPx/BQrL6kEQDUteWgDlblQUv+2cVSaL86FIkrwdwbHkaDopqVNK4FK+wRUARzM9NGkoQDhhGLDrhmsClWBsAkXZYX+oBGQ6xisbDpi1ZiFKmDDj5/QLMJ1fapBzk6DPvSjRceWv2HpRP280D5AIUYB0hXoyjhX8vv1rE8PHjdNoiTuKVXLgdBxqUHug6Y6y5VZOntOPUZTm+ecZZ5acdaVSp68r1Qtuk04lHiw6hNCzNpz8Vg==",
            "ctl00$hfServerTime": "1669525723096.74",
            "ctl00$ContentPlaceHolder1$ddlForeign": "1",
            "ctl00$ContentPlaceHolder1$txtIDNOorPatientID": "",
            "ctl00$ContentPlaceHolder1$dd1BirthM": "7",
            "ctl00$ContentPlaceHolder1$txtVerificationCode":"", 
            "ctl00$ContentPlaceHolder1$hidStatus": "",
            "ctl00$ContentPlaceHolder1$FstClick":"", 
            "ctl00$ContentPlaceHolder1$SecClick":"", 
            "ctl00$ContentPlaceHolder1$msg":"", 
            "ctl00$ContentPlaceHolder1$msgdata":"", 
            "ctl00$ContentPlaceHolder1$msgpage": "/OINetReg/OINetReg.Reg/Reg_NetReg.aspx",
        }

        self.payloadD = {
            "ctl00$ScriptManager1":"ctl00$ContentPlaceHolder1$UpdatePanel2|ctl00$ContentPlaceHolder1$dd1BirthM",
            "__EVENTTARGET":"ctl00$ContentPlaceHolder1$dd1BirthM",
            "__EVENTARGUMENT": "",
            "__LASTFOCUS":"", 
            "__VIEWSTATE": "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",
            "__VIEWSTATEGENERATOR": "C8D299FD",
            "__VIEWSTATEENCRYPTED": "",
            "__EVENTVALIDATION":"ayDMtuYQwCBau/+sb75iWxAxDu47YDkEHWwI+582V/XcuNOrIWfWY3uBkxCOTkgbAHI3pAKJyp7QapB/166GSV39wczVypzGg4yshF+Cdz65g41V3XCkRXWTNUTWOTOADFUo1f3tTnquhXwrOrkvsPx/BQrL6kEQDUteWgDlblQUv+2cVSaL86FIkrwdwbHkaDopqVNK4FK+wRUARzM9NGkoQDhhGLDrhmsClWBsAkXZYX+oBGQ6xisbDpi1ZiFKmDDj5/QLMJ1fapBzk6DPvSjRceWv2HpRP280D5AIUYB0hXoyjhX8vv1rE8PHjdNoiTuKVXLgdBxqUHug6Y6y5VZOntOPUZTm+ecZZ5acdaVSp68r1Qtuk04lHiw6hNCzNpz8Vg==",
            "ctl00$hfServerTime": "1669525723096.74",
            "ctl00$ContentPlaceHolder1$ddlForeign": "1",
            "ctl00$ContentPlaceHolder1$txtIDNOorPatientID": "",
            "ctl00$ContentPlaceHolder1$dd1BirthM": "7",
            "ctl00$ContentPlaceHolder1$dd1BirthD": "1",
            "ctl00$ContentPlaceHolder1$txtVerificationCode":"", 
            "ctl00$ContentPlaceHolder1$hidStatus": "",
            "ctl00$ContentPlaceHolder1$FstClick":"", 
            "ctl00$ContentPlaceHolder1$SecClick":"", 
            "ctl00$ContentPlaceHolder1$msg":"", 
            "ctl00$ContentPlaceHolder1$msgdata":"", 
            "ctl00$ContentPlaceHolder1$msgpage": "/OINetReg/OINetReg.Reg/Reg_NetReg.aspx",
        }

        self.OK_Payload = {
            "ctl00$ScriptManager1": "ctl00$ContentPlaceHolder1$UpdatePanel1|ctl00$ContentPlaceHolder1$btnReg",
            "__EVENTTARGET": "" ,
            "__EVENTARGUMENT": "",
            "__LASTFOCUS": "",
            "__VIEWSTATE":"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",
            "__VIEWSTATEGENERATOR": "C8D299FD",
            "__VIEWSTATEENCRYPTED": "",
            "__EVENTVALIDATION": "o1fc/HVuBaFnjvyfoPnnJuyZNaRFq3ryjuV4eggF2cNESqpOVaYZK3h5OUuANbX4fEa7AL6UAL2qiMMK5BSzmYYk5DmsQ9ADfFX5BeEiiY1wxXK3Q4YEQlzEXAlINiMInf1IxVze6dkY0gjyaIBdpyluQAk92abVyDdmwSJwmXe5GD2YS4vLDSkvMj9AD6pCRes9FT+tEdP35tNtaLO5x+HojpgoQR7SwAGGMqLhimD1Wn0XahHcGFm6wp/d5J0Wl73pM/jrxhFq4ISbjNzJBI+KQbTdEW6ziUUs/5f2DvAA0IzLw9C5135JEkjQX8a+Q2bfBCOoqs2NT1poLA6Rrikx77MjotuRTKR+cYAlgSjXEvR/VfvbOvDa2xS3Oojh2QJa53TcnYcNz2x53SBxM5YJ66vCL+kMxYSQnWTwNQaUwMZmlGTMLrTnNe4OSfPl+/p/fbnDwj5eyV6l88YjsXU54ARZbpUbV3HsSUYjMFD+347Bxi3xKyfp4Mku/mnw9DqQpe/cCetuPvvA2SmNGq/5bFf5aPYM/BYfqd0dy/PUrzhsdy2l0CyX3fpjvAw/8hZSeQmAzzJUhTCYLjrPJ76p5mH1nf92tsma8wjWcwCyK9o0DP/9WwpJTY5DctecPVsrEA==",
            "ctl00$hfServerTime": "1669525723096.74",
            "ctl00$ContentPlaceHolder1$ddlForeign": "1",
            "ctl00$ContentPlaceHolder1$txtIDNOorPatientID": "H125087083",
            "ctl00$ContentPlaceHolder1$dd1BirthM": "7",
            "ctl00$ContentPlaceHolder1$dd1BirthD": "1",
            "ctl00$ContentPlaceHolder1$txtVerificationCode": "51492",
            "ctl00$ContentPlaceHolder1$hidStatus": "",
            "ctl00$ContentPlaceHolder1$FstClick": "",
            "ctl00$ContentPlaceHolder1$SecClick": "",
            "ctl00$ContentPlaceHolder1$msg": "",
            "ctl00$ContentPlaceHolder1$msgdata":"", 
            "ctl00$ContentPlaceHolder1$msgpage": "/OINetReg/OINetReg.Reg/Reg_NetReg.aspx",
            "ctl00$ContentPlaceHolder1$btnReg": "查詢",
        }
        self.browser = browser

    def run(self):
        while True:
            if self._PDFData() and self.window.RunStatus:
                for persionData in self.Data :
                    print(persionData)
                    if (self.currentNum <= self.EndNum) and (self.currentPage <= self.EndPage) and self.window.RunStatus:
                        content = "姓名 : " + persionData['Name'] + "\n身分證字號 : " + persionData['ID'] + "\n出生日期 : " + persionData['Born'] + "\n查詢醫院 : 部立台中醫院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(persionData['Name'], persionData['ID'], persionData['Born'].split('/')[0],persionData['Born'].split('/')[1],persionData['Born'].split('/')[2])
                        self._startBrowser(persionData['Name'],persionData['ID'])
                        time.sleep(2)
                        self.currentNum += 1
                    else:
                        break
                self.currentNum = 1
                self.currentPage += 1
            else:
                self.window.setStatusText(content="~比對完成~",x=0.35,y=0.7,size=24)
                self.window.GUIRestart()
                self._endBrowser()
                break
        del self

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str):
        self.OK_Payload['ctl00$ContentPlaceHolder1$txtIDNOorPatientID'] = ID
        self.OK_Payload['ctl00$ContentPlaceHolder1$dd1BirthM'] = str(int(month))
        self.OK_Payload['ctl00$ContentPlaceHolder1$dd1BirthD'] = str(int(day))
        with httpx.Client(http2=True) as client :
            respone = client.get("https://www03.taic.mohw.gov.tw/OINetReg/OINetReg.Reg/Reg_RegConfirm.aspx")
            soup = BeautifulSoup(respone.content,"html.parser")
            
            # 發送月份請求
            self.payloadM["ctl00$ScriptManager1"] = "ctl00$ContentPlaceHolder1$UpdatePanel2|ctl00$ContentPlaceHolder1$dd1BirthM"
            self.payloadM["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
            self.payloadM["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
            self.payloadM["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
            self.payloadM["ctl00$hfServerTime"] = soup.find("input",{"id":"ctl00_hfServerTime"}).get("value")
            self.payloadM["ctl00$ContentPlaceHolder1$dd1BirthM"] = "7"
            respone = client.post("https://www03.taic.mohw.gov.tw/OINetReg/OINetReg.Reg/Reg_RegConfirm.aspx",data=self.payloadM,headers=self.header)
            soup = BeautifulSoup(respone.content,"html.parser")
            
            # 發送日期請求
            self.payloadD["ctl00$ScriptManager1"] = "ctl00$ContentPlaceHolder1$UpdatePanel2|ctl00$ContentPlaceHolder1$dd1BirthD"
            self.payloadD["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
            self.payloadD["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
            self.payloadD["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
            self.payloadD["ctl00$hfServerTime"] = soup.find("input",{"id":"ctl00_hfServerTime"}).get("value")
            self.payloadD["ctl00$ContentPlaceHolder1$dd1BirthD"] = "1"
            respone = client.post("https://www03.taic.mohw.gov.tw/OINetReg/OINetReg.Reg/Reg_RegConfirm.aspx",data=self.payloadD,headers=self.header)
            soup = BeautifulSoup(respone.content,"html.parser")
            
            while True:
                # 請求驗證碼
                respone = client.get("https://www03.taic.mohw.gov.tw/OINetReg/OINetReg.Reg/VerificationCode.aspx")
                with open("VaildCode.png","wb") as f :
                    f.write(respone.content)

                # 發送正式請求
                self.OK_Payload["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
                self.OK_Payload["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
                self.OK_Payload["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
                self.OK_Payload["ctl00$hfServerTime"] = soup.find("input",{"id":"ctl00_hfServerTime"}).get("value")
                self.OK_Payload["ctl00$ContentPlaceHolder1$txtVerificationCode"] = self._ParseCaptcha()
                respone = client.post("https://www03.taic.mohw.gov.tw/OINetReg/OINetReg.Reg/Reg_RegConfirm.aspx",data=self.OK_Payload,headers=self.header)
                with open("test.html","wb") as f :
                    f.write(respone.content)
                    if not self._CKCaptcha(respone.content,"span","驗證碼錯誤! 請輸入正確的驗證碼！"):
                        with open("reslut.html","w",encoding="utf-8") as f :
                            f.write(self._changeHTMLStyle(respone.content,"https://www03.taic.mohw.gov.tw/OINetReg/",""))
                        time.sleep(2)
                        break
                    else:
                        self.window.setStatusText(content="驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                        time.sleep(1)
                        content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 林新醫院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("就診序號",(name + '_' + ID + '_部立台中醫院.png')) :
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

    def _PDFData(self) -> bool:
        # print("Current : " + str(self.currentPage) + "  End : " + str(self.EndPage))
        if (self.currentPage <= self.EndPage):
            mPDFReader = PDFReader(self.window,self.filePath)
            status, self.Data = mPDFReader.GetData(self.currentPage-1)
            return status
        else:
            return False
    
    def _changeHTMLStyle(self,page_content,targer1:str,targer2:str):
        soup = BeautifulSoup(page_content,"html.parser")
        # 替換屬性內容，強制複寫資源路徑(針對img)
        Tags = soup.find_all(['img'])
        for Tag in Tags:
            try:
                Tag['src'] = Tag['src'].replace('..',targer1) 
            except:
                pass
        # 替換屬性內容，強制複寫資源路徑(針對link)
        Tags = soup.find_all(['link'])
        for Tag in Tags:
            try:
                Tag['href'] = Tag['href'].replace('..',targer1)
            except:
                pass
        # 替換屬性內容，強制複寫資源路徑(針對input)
        Tags = soup.find_all(['input'])
        for Tag in Tags:
            try:
                Tag['src'] = Tag['src'].replace('..',targer1) 
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

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")



