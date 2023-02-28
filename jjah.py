import time
from bs4 import BeautifulSoup
import ddddocr
import os
from PDFReader import PDFReader
# 2022/12/24加入
from LogController import Log
from VPNClient import VPN
from VPNWindow import VPNWindow
from tkinter import messagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert


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
        self.maxError =10
        self.errorNum = 0
        self.url = "https://www.jah.org.tw/JCHReg/Query/J"
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
            if self.window.RunStatus:
                if self._PDFData(self.page):
                    for self.idx in range(self.currentNum-1,self.datalen) :
                        print(self.Data[self.idx])
                        if ((self.page != self.EndPage) and (self.idx != self.EndNum)) and self.window.RunStatus:
                            # 查詢狀態
                            Q_Status = False
                            # 初診查詢
                            content = "姓名 : " + self.Data[self.idx]['Name'] + "(初診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                            self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                            self._getReslut(self.Data[self.idx]['Name']+"(初診)", self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                            Q_Status = self._startBrowser(self.Data[self.idx]['Name'] + "(初診)",self.Data[self.idx]['ID'])
                            self.log.write(self.Data[self.idx]['Name'] + "(初診)",self.Data[self.idx]['ID'],"大里仁愛醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                            time.sleep(2)
                            # 複診查詢
                            if not(Q_Status) and self.window.RunStatus:
                                content = "姓名 : " + self.Data[self.idx]['Name'] + "(複診)\n身分證字號 : " + self.Data[self.idx]['ID'] + "\n出生日期 : " + self.Data[self.idx]['Born'] + "\n查詢醫院 : 大里仁愛醫院\n當前第" + str(self.page + 1) + "頁，第" + str(self.idx + 1) + "筆"
                                self.window.setStatusText(content=content,x=0.3,y=0.75,size=12)
                                Q_Status = self._getReslut(self.Data[self.idx]['Name']+"(複診)", self.Data[self.idx]['ID'], self.Data[self.idx]['Born'].split('/')[0],self.Data[self.idx]['Born'].split('/')[1],self.Data[self.idx]['Born'].split('/')[2])
                                self._startBrowser(self.Data[self.idx]['Name'],self.Data[self.idx]['ID'])
                                self.log.write(self.Data[self.idx]['Name'] + "(複診)",self.Data[self.idx]['ID'],"大里仁愛醫院",self.Data[self.idx]['Born'],str(self.page + 1),str(self.idx + 1))
                                time.sleep(2)
                            else:
                                break
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
        while True:
            try:
                self.browser.get(self.url)
                time.sleep(2)
                print(name)
                if "(初診)" in name:
                    self.browser.find_element(By.XPATH, '//*[@id="QueryForm"]/div[1]/label[2]/span').click()
                    time.sleep(2)
                print(1)
                # ActionChains(driver).move_to_element(driver.find_element_by_xpath('//*[@id="QueryForm"]/div[1]/label[2]/input')).click().perform()
                self.browser.find_element(By.XPATH, '//*[@id="QueryForm"]/div[2]/input').send_keys(ID)
                time.sleep(2)
                print(2)
                bornDate = str(int(year) + 1911) + month + day
                self.browser.find_element(By.XPATH, '//*[@id="birthday"]').send_keys(bornDate)
                time.sleep(2)
                print(3)
                Captcha = self._ParseCaptcha4Img(self.browser.find_element(By.XPATH, '//*[@id="captcha"]'))
                time.sleep(3)
                print(4)
                self.browser.find_element(By.XPATH, '//*[@id="QueryForm"]/div[2]/div[3]/input').send_keys(Captcha)
                time.sleep(3)
                print(5)
                self.browser.find_element(By.XPATH, '//*[@id="QueryForm"]/div[2]/div[4]/div/button[1]').click()
                time.sleep(3)
                print(6)
                dropDown="var q=document.documentElement.scrollTop=500"  
                self.browser.execute_script(dropDown)
                time.sleep(3)
                print(7)
                reslut = self._CKCaptcha("Web", "驗證碼比對錯誤，請重新輸入")
                if not reslut:
                    print("8-1")
                    self.window.setStatusText(content="因驗證碼錯誤，系統正重新查詢",x=0.2,y=0.7,size=20)
                else:
                    dropUp="var q=document.documentElement.scrollTop=0"  
                    self.browser.execute_script(dropUp)
                    print("8-2")
                    self.errorNum = 0
                    break
            except:
                print("發生錯誤即將重試(" + str(self.errorNum) + ")")
                self.window.setStatusText(content="~發生錯誤(" + str(self.errorNum) + ")，準備再次嘗試~\n~等候重新執行當前人員查詢~",x=0.3,y=0.8,size=12)
                if(self.errorNum >= self.maxError):
                    messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
                    os._exit(0)
                self.errorNum += 1
                time.sleep(5)
        
    def _startBrowser(self,name,ID) -> bool:
        if self._Screenshot("取消掛號",(name + '_' + ID + '_大里仁愛醫院.png')) :
            self.window.setStatusText(content="~條件符合，已截圖保存~",x=0.25,y=0.7,size=24)
            return True
        else:
            self.window.setStatusText(content="~不符合截圖標準~",x=0.3,y=0.7,size=24)
            return False

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
    
    # 檢查驗證是否成功
    def _CKCaptcha(self, CK_Method, msg):
        try:
            if CK_Method == "Alert":
                print("這是彈窗訊息 : " + Alert(self.browser).text.strip())
                print("這是條件訊息 : " + msg)
                if Alert(self.browser).text.strip() == msg:
                    Alert(self.browser).accept()
                    return False
                else:
                    print(Alert(self.browser).text)
                    Alert(self.browser).accept()
                    return True
            if CK_Method == "Web":
                html = str(BeautifulSoup(self.browser.page_source,"html.parser"))
                if msg in html:
                    return False
                return True
        except Exception as e :
            print("這是錯誤訊息 : " + str(e))
            return True
    
    
    # 驗證碼辨識
    def _ParseCaptcha4Img(self, element):
        image_base64 = self.browser.execute_script("\
                var ele = arguments[0];\
                var cnv = document.createElement('canvas');\
                cnv.width = ele.width;\
                cnv.height = ele.height;\
                cnv.fillStyle = '#FFFFFF';\
                cnv.getContext('2d').drawImage(ele,0,0);\
                return cnv.toDataURL('image/png').substring(22);\
            ",element)
        orc = ddddocr.DdddOcr()
        result = orc.classification(image_base64)
        return result

    def _endBrowser(self):
        self.browser.quit()

    def __del__(self):
        print("物件刪除")
