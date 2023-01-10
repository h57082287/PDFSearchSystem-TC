import httpx
from bs4 import BeautifulSoup
import ddddocr
import os
import time
from PDFReader import PDFReader
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox

# 童綜合醫院
class RG():
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
        self.page = 0
        self.datalen = 0
        self.log = Log()

        # 建立header
        self.headers = {
                    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.47'
                }
        # 建立payload
        self.payload = {
                        '__EVENTTARGET': '',
                        '__EVENTARGUMENT': '',
                        '__VIEWSTATE': '/wEPDwUKLTM1ODk1MTQwMg9kFgJmD2QWAgIFD2QWAgIBD2QWAgIVDzwrAAsAZGTjtjGPHsFF6aiKkbzNalc9WkBmO0uYT3oK4jaSH3FgrQ==',
                        '__VIEWSTATEGENERATOR': 'C8064417',
                        '__EVENTVALIDATION': '/wEdADCntmjCPnemxwEgA8Qw39TO+F+XK/o0ZXRunx0e+5mTEkOmLsViywxU1fjDNw6LeeyNLCVxVr7ocYBfPqFwSsrKKmSs05+DODKdILgmyiIAp6KtLfallEXN7ARHVTyRJMWsWvsuLGrBj/e6Daa8Ptp+eLVgCewbRHI4cJKaDjhMyLUq80sN4EnrYMuPMu+DpFb65hNpAzHM27mb9stByOsy2tNKRK7Af8mPDhoUzZ6+DkwkR2cGWMaAUEp07yvnxd0Fns2NIoQw54TQMBHdf58t1fsEQpc97cM/MQ85BoFkQYxl6kHlttWfQU0UCigQEN9a20QvoFomklTujOoZVHgdRggsczUkN7CErJIIFb4oIAv1JL2J79aJstLMiETbgAz7HjPKnKtR/qrewkLjy63ZJ1gMI0uchWVQ10N83zkdI0k77J3/lvXamijpJmPBIc5SUxve6MtH2pC352mSvW4Ob7Tqy+Da37ayHMGn0HS+heikZ1IzjrIrwElvOhwsc/6d7lRda3Glht9iphnHvucqaKPix+d7sT1bNpeK7hsfAERPvyrrFB0I1IzoxCzRmYrrr+Ba2QETUIqKk6pK9g4Ud7TsaJns+AaEW6mRcGrRu71lVzi1AVNpMR5ewdirqQa88lVqSjsvZiLnYH+BCtszSwkJjvHggkBvFp+JdTaNbQf04z6qCchhqabhGZttcIOGcOoYCWfPut244Hc3CNy/COwy84sx7e+BrUFQbJMDB136eH8u5VFKwkcy+rpQs6jheaOD+sxh8SlbYLbODh6Oq8YTsIfuAZJmcd4kD6AKrbI44oxCKoRUj9FOf4y/Pwv4TsNCjeeW2nfOLtPmYVDtL3SeUlIdGk93/MP8QKHw06Ge3rHN2hIhaUkN/vXr4RZrny8l9V2zKkeI6YIFNAqu+OOW+jjiPfTFv8FDx5JtJqzX1PBLqZt/Bt7Uv3C7sxFws7p4ngo1V8u5xaLpj9akIJcsxFyJ/Hl60Z83VtrpigeUDhMOtfKrXL/ULF5Yz7CuPok95d+tZ758wraPnIbxjUG3Wf1wynmSysX4Llc6yw==',
                        'ctl00$ContentPlaceHolder2$tbID': 'H125087083',
                        'ctl00$ContentPlaceHolder2$btConfirm': '確認',
                        'ctl00$ContentPlaceHolder2$ddlMonth': '7',
                        'ctl00$ContentPlaceHolder2$ddlDay': '1',
                        'ctl00$ContentPlaceHolder2$TextBox1': '5376'
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
            if self._PDFData(self.page) and self.window.RunStatus:
                for self.idx in range(self.currentNum-1,self.datalen) :
                    print(self.Data[self.idx])
                    if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                        content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 童綜合醫院醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                        self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                        self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"童綜合醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
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
        self.payload['ctl00$ContentPlaceHolder2$tbID'] = ID
        self.payload['ctl00$ContentPlaceHolder2$ddlMonth'] = int(month)
        self.payload['ctl00$ContentPlaceHolder2$ddlDay'] = int(day)
        while True : # 加入VPN切換功能
            try:
                with httpx.Client(http2=True) as client :
                    # 進入網頁
                    respone = client.get('https://rg.sltung.com.tw/frmRgQuery.aspx?chartNam=' + ID)
                    soup = BeautifulSoup(respone.content,"html.parser")
                    
                    # 獲取隱藏驗證
                    self.payload['__VIEWSTATE'] = soup.find('input',{'id':'__VIEWSTATE'}).get('value')
                    self.payload['__VIEWSTATEGENERATOR'] = soup.find('input',{'id':'__VIEWSTATEGENERATOR'}).get('value')
                    self.payload['__EVENTVALIDATION'] = soup.find('input',{'id':'__EVENTVALIDATION'}).get('value')
                    
                    while True:
                        # 請求驗證碼
                        respone = client.get('https://rg.sltung.com.tw/img.aspx')
                        with open('VaildCode.png','wb') as f :
                            f.write(respone.content)
                        self.payload['ctl00$ContentPlaceHolder2$TextBox1'] = self._ParseCaptcha()

                        # 發送查詢請求
                        respone = client.post('https://rg.sltung.com.tw/frmRgQuery.aspx?chartNam=' + ID,data=(self.payload),headers=self.headers)
                        if not self._CKCaptcha(respone.content,"scrtpt","識別碼錯誤"):
                            with open("reslut.html",'wb') as f :
                                f.write(respone.content)
                            break
                        else:
                            self.window.setStatusText(content="驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                            time.sleep(1)
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 童綜合醫院醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                    break
            except httpx.ReadTimeout :
                try:
                    self.VPN.startVPN()
                    content = "姓名 : " + self.Data[self.idx]['Name'] + "\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 童綜合醫院醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                    self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break

    def _startBrowser(self,name,ID):
        self.browser.get(r'file:///' + os.path.dirname(os.path.abspath(__file__)) + '/reslut.html')
        if self._Screenshot("看診進度查詢",(name + '_' + ID + '_童綜合醫院.png')) :
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