import os 
import time
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup
import tkinter as tk
from VPNWindow import VPNWindow

# 規則說明 :
# 1. 各醫院物件需加入self.ChangeIPNow變數
# 2. 於App.py檔中加入VPN物件
# 3. 傳入MainWindows物件及各個醫院物件(App.py)
# 4. 各醫院新增ChangingIPCK funaction
# 5. 各醫院地的respone改成self.respone

class VPN():
    def __init__(self,MainWindowsObj,HospitalObj) -> None:
        self.window = MainWindowsObj
        self.hospital = HospitalObj
        self.Text_Rule = ["對不起!此ip查詢或取消資料次數過多;"]
        self.Code_Rule = [400,408,500,407,503,404]
        self.ovpnDir = ""
        self.fileList = ["",""]
        self.index = 0

        if (self.window.checkVal_AUVPNM.get()):
            self.hospital.ChangeIPNow = True
            VPNWindow(self)
            self.hospital.ChangeIPNow = False
        
    def _InstallationCkeck(self):
        cmd = 'start openvpn --config "C:\\Users\\h5708\\Downloads\\Surfshark_Config\\89.187.178.94_tcp.ovpn" --auth-user-pass login.conf'
        os.system(cmd)
        time.sleep(5)
        if self._CKVPNStatus:
            self._stopVPN()
        else:
            pass

    def _startVPN(self):
        if (self.index != len(self.fileList)) and self.window.checkVal_AUVPNM.get():
            cmd = 'start openvpn --config "' + self.fileList[self.index] + '" --auth-user-pass login.conf'
            # cmd = 'start openvpn --config "C:\\Users\\h5708\\Downloads\\Surfshark_Config\\89.187.178.94_tcp.ovpn" --auth-user-pass login.conf'
            os.system(cmd)
            self.index += 1
            self._memberControl()
        else:
            # self.window.RunStatus = False
            messagebox.showerror(title="IP替換失敗",message="所有可用ip已用盡或您本次並未啟用VPN功能!!!")
            self._stopVPN()
            os._exit(0)
    
    def watch_dog(self):
        while True:
            if(not self.window.RunStatus):
                self._stopVPN()
                break
            if((self.hospital.ChangeIPNow) or (self._ContentCK())) and self.window.RunStatus:
                self.window.setStatusText(content="檢測到IP異常，切換IP...",x=0.23,y=0.7,size=24)
                if(self._CKVPNStatus):
                    self._stopVPN()
                self._startVPN()
                self._CKIPChanged()
                self.hospital.ChangeIPNow = False
    
    def _ContentCK(self):
        if self.hospital.respone != None:
            for detail in self.Text_Rule:
                if detail in self.hospital.respone.text:
                    return True
            if self.hospital.respone.status_code in self.Code_Rule:
                return True
        return False
    
    def _memberControl(self):
        self.hospital.idx -= 1
        if self.hospital.idx <= 0 :
            self.hospital.currentPage -= 1
            self.hospital.idx = self.hospital.olddatalen -1
            self.hospital._PDFData()
        if self.hospital.currentPage <= 0 :
            self.hospital.currentPage = 0
            self.hospital.idx = 0
            self.hospital._PDFData()

    def _CKIPChanged(self):
        while True:
            try:
                time.sleep(5)
                respone = requests.get("https://myip.com.tw/")
                soup = BeautifulSoup(respone.content,"html.parser")
                self.window.setStatusText(content="當前IP :" + soup.find("font").text,x=0.23,y=0.7,size=24)
                time.sleep(3)
                break
            except:
                pass

    def _CKVPNStatus(self):
        cmd = "tasklist"
        Tasks = os.popen(cmd).read()
        return ("openvpn.exe" in Tasks)

    def _stopVPN(self):
        os.system("taskkill /F /IM openvpn.exe")
    
    # def ChangingIPCK(self):
    #     while(not self.IPStatus):
    #         pass
    #     self.IPStatus = True

    