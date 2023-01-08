import httpx
from bs4 import BeautifulSoup
import ddddocr
import os
import time
from PDFReader import PDFReader
from LogController import Log
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
import requests

# 林新醫院
class LSHOSP():
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
        self.header = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36"
        }

        self.payloadM = {
            "__EVENTTARGET": "ctl00$ContentPlaceHolder1$dd1BirthM",
            "__EVENTARGUMENT": "",
            "__LASTFOCUS": "",
            "__VIEWSTATE": "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",
            "__VIEWSTATEGENERATOR": "A576A846",
            "__VIEWSTATEENCRYPTED": "",
            "__EVENTVALIDATION": "e9fdkLqRF0GL9cZWPUPW9YdGFwaZv+70544YYyIr94MKCgfYLflY3axNSIk8xzFLy0WX5zO1iww01mf5bYd/D+KKKAwe2gzZWOj548iq0XCI+odEbgzkd45Ok+eCNfuhapBnRGWXgMIo3/Yl+TkP+ZCUN6XcltUijSh6265EovOrVSX1JKKwzdcCP24Ozt8+HreelWhPj29Z8ZFAbThoD25hbrBSEIZP9d6ywlPaElPimPmX+c+ANqVpOfOhZ2mH3YO6JjROyZefFeeYYgdNe5ox/GvKVWCc3QB4CCFoIS7YlQUr3D6ImtqnUI/o2l5UnvvyXlqNi4SvQp0/2Qw3Gi2kdi0=",
            "ctl00$hfServerTime": "False",
            "ctl00$ContentPlaceHolder1$ddlForeign": "1",
            "ctl00$ContentPlaceHolder1$txtIDNOorPatientID": "",
            "ctl00$ContentPlaceHolder1$dd1BirthM": "7",
            "ctl00$ContentPlaceHolder1$txtVerificationCode": "",
            "ctl00$ContentPlaceHolder1$hidStatus": "",
            "ctl00$ContentPlaceHolder1$FstClick": "",
            "ctl00$ContentPlaceHolder1$SecClick": "",
        }

        self.payloadD = {
            "__EVENTTARGET": "ctl00$ContentPlaceHolder1$dd1BirthM",
            "__EVENTARGUMENT": "",
            "__LASTFOCUS": "",
            "__VIEWSTATE": "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",
            "__VIEWSTATEGENERATOR": "A576A846",
            "__VIEWSTATEENCRYPTED": "",
            "__EVENTVALIDATION": "e9fdkLqRF0GL9cZWPUPW9YdGFwaZv+70544YYyIr94MKCgfYLflY3axNSIk8xzFLy0WX5zO1iww01mf5bYd/D+KKKAwe2gzZWOj548iq0XCI+odEbgzkd45Ok+eCNfuhapBnRGWXgMIo3/Yl+TkP+ZCUN6XcltUijSh6265EovOrVSX1JKKwzdcCP24Ozt8+HreelWhPj29Z8ZFAbThoD25hbrBSEIZP9d6ywlPaElPimPmX+c+ANqVpOfOhZ2mH3YO6JjROyZefFeeYYgdNe5ox/GvKVWCc3QB4CCFoIS7YlQUr3D6ImtqnUI/o2l5UnvvyXlqNi4SvQp0/2Qw3Gi2kdi0=",
            "ctl00$hfServerTime": "False",
            "ctl00$ContentPlaceHolder1$ddlForeign": "1",
            "ctl00$ContentPlaceHolder1$txtIDNOorPatientID": "",
            "ctl00$ContentPlaceHolder1$dd1BirthM": "7",
            "ctl00$ContentPlaceHolder1$dd1BirthD" : "1",
            "ctl00$ContentPlaceHolder1$txtVerificationCode": "",
            "ctl00$ContentPlaceHolder1$hidStatus": "",
            "ctl00$ContentPlaceHolder1$FstClick": "",
            "ctl00$ContentPlaceHolder1$SecClick": "",
        }

        self.OK_Payload = {
            "__EVENTTARGET": "",
            "__EVENTARGUMENT":"" ,
            "__LASTFOCUS": "",
            "__VIEWSTATE": "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",
            "__VIEWSTATEGENERATOR":"A576A846",
            "__VIEWSTATEENCRYPTED": "",
            "__EVENTVALIDATION": "m2WsxKITmjvcb6jYuf9tyeD8drExM3KxLyJk/TISuYYf7yMeFOhjCBgPJjbiQiGVMGFWsWCmpTF8HrQ/BBCxyQAA6NGhHMfDMvFj1kGmrulIMMKTpItcgw6qGY8ZZWpPy7Mg4116kN2L1ayaiaS+JDpWBopgYoFaaKhX+0Bxw3eINyXXlx5+eFp90/ASXoDC6dTBH4dkw8y3BOKsGAcxZjIwtocEjCzAvikA/awSzJVTrKnr8OoWBXVzKF8snA7CfFgLE7Od5IlhoNespHvubKBrg3UiaVsifUNYSnMN53aCcDqGVC0NuMwFGgJb6nkjlQIfzHC7gD8AcyZAqdU/htHsmu4dwvpojy8aJRnxYcg6FxzwalLWo5UOoBS+6dGPtJEvG5bISFXXtoT9TCSbgxQ+k51PkJZMiQJrbRxVBJoavYu/JGwCQnqIaLxMa+KIEMEsDg773X3/WUQ+4RkIukJDc7S8mYq3OT1i37+0Ko/EQTwhbf2iW08e/UWAqyhgttXisaG0lx+r3HoQmVtunXBlAE0hu9237kmPkEgfBTzMgpW8jY70SXULTH6MRFp1ieUseRRMDk4eQD31qIcXPONY2Lo=",
            "ctl00$hfServerTime": "1669531076000",
            "ctl00$ContentPlaceHolder1$ddlForeign": "1",
            "ctl00$ContentPlaceHolder1$txtIDNOorPatientID": "H125087083",
            "ctl00$ContentPlaceHolder1$dd1BirthM": "7",
            "ctl00$ContentPlaceHolder1$dd1BirthD": "1",
            "ctl00$ContentPlaceHolder1$txtVerificationCode": "MGGmtfy",
            "ctl00$ContentPlaceHolder1$hidStatus": "",
            "ctl00$ContentPlaceHolder1$btnReg": "查詢",
            "ctl00$ContentPlaceHolder1$FstClick": "",
            "ctl00$ContentPlaceHolder1$SecClick": "",
        }
        self.browser = browser
        # 各醫院新增項目
        self.respone = None
        self.ChangeIPNow = False
        self.idx = 0
        self.datalen = 0
        self.olddatalen = 0
        self.log = Log()

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
            if self._PDFData(self.page) and self.window.RunStatus:
                for self.idx in range(self.currentNum-1,self.datalen) :
                    print(self.Data[self.idx])
                    if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                        content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 林新醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                        self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                        self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"林新醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
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
        self.OK_Payload['ctl00$ContentPlaceHolder1$txtIDNOorPatientID'] = ID
        self.OK_Payload['ctl00$ContentPlaceHolder1$dd1BirthM'] = str(int(month))
        self.OK_Payload['ctl00$ContentPlaceHolder1$dd1BirthD'] = str(int(day))
        while True:
            try:
                with httpx.Client(http2=True) as client :
                    self.respone = client.get("http://www.lshosp.com.tw:8001/OINetReg/OINetReg.Reg/Reg_RegConfirm1.aspx")
                    soup = BeautifulSoup(self.respone.content,"html.parser")
                    
                    # 發送月份請求
                    self.payloadM["__EVENTTARGET"] = "ctl00$ContentPlaceHolder1$dd1BirthM"
                    self.payloadM["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
                    self.payloadM["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
                    self.payloadM["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
                    self.payloadM["ctl00$hfServerTime"] = soup.find("input",{"id":"ctl00_hfServerTime"}).get("value")
                    self.payloadM["ctl00$ContentPlaceHolder1$dd1BirthM"] = str(int(month))
                    self.respone = client.post("http://www.lshosp.com.tw:8001/OINetReg/OINetReg.Reg/Reg_RegConfirm1.aspx",data=self.payloadM,headers=self.header)
                    soup = BeautifulSoup(self.respone.content,"html.parser")
                    
                    # 發送日期請求
                    self.payloadD["__EVENTTARGET"] = "ctl00$ContentPlaceHolder1$dd1BirthD"
                    self.payloadD["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
                    self.payloadD["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
                    self.payloadD["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
                    self.payloadD["ctl00$hfServerTime"] = soup.find("input",{"id":"ctl00_hfServerTime"}).get("value")
                    self.payloadD["ctl00$ContentPlaceHolder1$dd1BirthM"] = str(int(month))
                    self.payloadD["ctl00$ContentPlaceHolder1$dd1BirthD"] = str(int(day))
                    self.respone = client.post("http://www.lshosp.com.tw:8001/OINetReg/OINetReg.Reg/Reg_RegConfirm1.aspx",data=self.payloadD,headers=self.header)
                    soup = BeautifulSoup(self.respone.content,"html.parser")
                    
                    while True :
                        # 請求驗證碼
                        self.respone = client.get("http://www.lshosp.com.tw:8001/OINetReg/OINetReg.Reg/ValidateNumber.ashx")
                        with open("VaildCode.png","wb") as f :
                            f.write(self.respone.content)

                        # 發送正式請求
                        self.OK_Payload["__VIEWSTATE"] = soup.find("input",{"id":"__VIEWSTATE"}).get("value")
                        self.OK_Payload["__VIEWSTATEGENERATOR"] = soup.find("input",{"id":"__VIEWSTATEGENERATOR"}).get("value")
                        self.OK_Payload["__EVENTVALIDATION"] = soup.find("input",{"id":"__EVENTVALIDATION"}).get("value")
                        self.OK_Payload["ctl00$hfServerTime"] = soup.find("input",{"id":"ctl00_hfServerTime"}).get("value")
                        self.OK_Payload["ctl00$ContentPlaceHolder1$txtVerificationCode"] = self._ParseCaptcha()
                        self.respone = client.post("http://www.lshosp.com.tw:8001/OINetReg/OINetReg.Reg/Reg_RegConfirm1.aspx",data=self.OK_Payload,headers=self.header)
                        if not self._CKCaptcha(self.respone.content,"span","驗證碼錯誤! 請輸入正確的驗證碼！"):
                            with open("reslut.html","w",encoding="utf-8") as f :
                                f.write(self._changeHTMLStyle(self.respone.content,"http://www.lshosp.com.tw:8001/OINetReg","http://www.lshosp.com.tw"))
                            time.sleep(2)
                            break
                        else:
                            self.window.setStatusText(content="驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                            time.sleep(1)
                            content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 林新醫院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            time.sleep(2)
                break
            except requests.exceptions.ConnectTimeout:
                print("發生時間例外")
                try:
                    self.VPN.startVPN()
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break
            except AttributeError:
                print("發生找不到內容例外")
                try:
                    self.VPN.startVPN()
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break



    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("就診序號",(name + '_' + ID + '_林新醫院.png')) :
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

    # 各醫院新增項目
    def _ChangingIPCK(self):
        while(self.ChangeIPNow):
            pass
        self.ChangeIPNow = False

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")