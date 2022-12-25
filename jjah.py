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


        # 建立header
        self.headers = {
                'accept': 'application/json, text/javascript, */*; q=0.01',
                'accept-encoding': 'gzip, deflate, br',
                'accept-language': 'zh-TW,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
                'content-length': '746',
                'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
                # 'cookie': '_ga=GA1.3.1523170946.1666280151; _gid=GA1.3.1436510728.1666280151; __RequestVerificationToken_L0pDSFJlZw2=-mGGcLfgt3WEuJX09Zxe9eo4KX4ILfHugZQxLaU998AIpR6yV63JoxQpBDNYrJf-k34T8YfZXiJPE2UpML6sGqNQ1tsgFPhPzeOOnltezMM1; ASP.NET_SessionId=omz1s1v0otiemuzwq2mvxr13'
                'dnt': '1',
                'origin': 'https://www.jah.org.tw',
                'referer': 'https://www.jah.org.tw/JCHReg/Query/U',
                'sec-ch-ua': '"Chromium";v="106", "Microsoft Edge";v="106", "Not;A=Brand";v="99"',
                'sec-ch-ua-mobile': '?0',
                'sec-ch-ua-platform': '"Windows"',
                'sec-fetch-dest': 'empty',
                'sec-fetch-mode': 'cors',
                'sec-fetch-site': 'same-origin',
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.47',
                'x-requested-with': 'XMLHttpRequest'
            }
        # 建立payload
        self.payload = {
                '__RequestVerificationToken': 'QGThtpG0yW4bCw0X3ZkdtYHZF7tH4BDr7JWzQT9YIdABsYzOEm_eNPoF8FHFUhQzCsDS2kZm2OaAJO5W65Ou7X2lc56Dyeq39MFJcAkZSDc1',
                'Func': 'QuerySearch',
                'Val': ''
            }
        # value資料
        self.Val = [
                {"name":"hospitalID","value":"J"},
                {"name":"patNumber","value":""},
                {"name":"isFirst","value":"Y"},
                {"name":"idType","value":"2"},
                {"name":"idNumber","value":"H125087083"},
                {"name":"birthday","value":"19980701"},
                {"name":"verification","value":"j4rx"}
            ]
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
            if self._PDFData(self.page) and self.window.RunStatus:
                for self.idx in range(self.currentNum-1,self.datalen) :
                    print(self.Data[self.idx])
                    if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                        # 查詢狀態
                        Q_Status = False
                        # 初診查詢
                        self.Val[2]['value'] = 'Y'
                        content = "姓名 : " + self.Data[self.idx]['Name'] + "(初診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                        self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                        Q_Status = self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                        self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                        self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"大里仁愛醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                        time.sleep(2)
                        # 複診查詢
                        if not(Q_Status) and self.window.RunStatus:
                            self.Val[2]['value'] = 'N'
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "(複診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            Q_Status = self._getReslut(self.Data[self.idx]['Name'], self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'],"大里仁愛醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            time.sleep(2)
                        else:
                            break
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

    def _getReslut(self,name:str, ID:str, year:str, month:str, day:str) -> bool:
        self.Val[4]['value'] = ID
        self.Val[5]['value'] = year + month + day
        status = False
        while True:
            try:
                with httpx.Client(http2=True) as client :
                    respone = client.get('https://www.jah.org.tw/JCHReg/Query/U')
                    soup = BeautifulSoup(respone.content,"html.parser")
                    time.sleep(1)

                    # 獲取隱藏元素
                    self.payload['__RequestVerificationToken'] = soup.find('form',{'id':'QueryForm'}).find('input',{'name':'__RequestVerificationToken'}).get('value')

                    # 利用迴圈自動重試
                    while True:
                        # 請求驗證碼
                        respone = client.get('https://www.jah.org.tw/JCHReg/Content/BuildCaptcha.aspx')
                        with open('VaildCode.png','wb') as f :
                            f.write(respone.content)
                        self.Val[6]['value'] = self._ParseCaptcha()

                        # 發送請求
                        self.payload['Val'] = quote(str(self.Val))
                        time.sleep(5)
                        respone = client.post('https://www.jah.org.tw/JCHReg/Ajax',data=self.payload)
                        if respone.json()['QueryList'] != '' :
                            status = True
                        if self._JSONDataToHTML(respone.json()) :
                            break
                        else:
                            self.window.setStatusText(content="驗證碼錯誤，系統正重新查詢",x=0.2,y=0.8,size=20)
                            time.sleep(1)
                            content = "姓名 : " + name + "\n身分證字號 : " + ID + "\n出生日期 : " + (year + "/" + month + "/" + day) + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.currentPage) + "頁，第" + str(self.currentNum) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                    break
            except requests.exceptions.ConnectTimeout:
                try:
                    self.VPN.startVPN()
                except:
                    messagebox.showerror("啟動VPN發生錯誤","無法啟動VPN輪轉功能，可能是您並未於設定裡允許'啟動VPN'的功能")
                    self.window.Runstatus = False
                    break
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
