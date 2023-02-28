import os 
import time
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup

# 規則說明 :
# 1. 各醫院物件需加入self.ChangeIPNow變數
# 2. 於App.py檔中加入VPN物件
# 3. 傳入MainWindows物件及各個醫院物件(App.py)
# 4. 各醫院新增ChangingIPCK funaction
# 5. 各醫院地的respone改成self.respone

class VPN():
    def __init__(self,MainWindowsObj) -> None:
        self.window = MainWindowsObj
        self.ovpnDir = ""
        self.fileList = []
        self.index = self._VPNHistory("r")
        
    def InstallationCkeck(self):
        self.window.setStatusText(content="檢測OpenVPN安裝狀態...",x=0.23,y=0.7,size=24)
        cmd = 'start openvpn --config "' + self.ovpnDir + '/' + self.fileList[self.index] + '" --auth-user-pass login.conf'
        os.system(cmd)
        time.sleep(20)
        if self._CKVPNStatus():
            self.window.setStatusText(content="已檢測到OpenVPN環境",x=0.23,y=0.7,size=24)
            self.stopVPN()
            return True
        else:
            return False

    def startVPN(self):
        self._VPNHistory("w", self.index)
        print("VPN-1")
        if (self.index != len(self.fileList)) and self.window.checkVal_AUVPNM.get():
            print("VPN-2")
            if self._CKVPNStatus():
                print("VPN-3")
                self.stopVPN()
            cmd = 'start openvpn --config "' + self.ovpnDir + '/' +self.fileList[self.index] + '" --auth-user-pass login.conf'
            os.system(cmd)
            print("VPN-4")
            self.index += 1
            print("VPN-5")
            self._CKIPChanged()
            print("VPN-6")
        else:
            print("無法替換ip")
            self.window.RunStatus = False
            messagebox.showerror(title="IP替換失敗",message="所有可用ip已用盡或您本次並未啟用VPN功能!!!")
            os.remove("VPN.conf")
            self.stopVPN()
            os._exit(0)
    
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
    
    def _getFileList(self):
        for file in os.listdir(self.ovpnDir):
            if "tcp" in file :
                self.fileList.append(file)

    def _CKVPNStatus(self):
        cmd = "tasklist"
        Tasks = os.popen(cmd).read()
        return ("openvpn.exe" in Tasks)

    def stopVPN(self):
        os.system("taskkill /F /IM openvpn.exe")
    
    def _VPNHistory(self, mode:str, data=None):
        if mode == "w":
            try:
                os.remove("VPN.conf")
            except:
                pass
            with open("VPN.conf", "a+") as f:
                print(data)
                f.writelines(str(data))
        elif mode == "r":
            with open("VPN.conf", "a+") as f:
                try:
                    return int(f.readline())
                except:
                    return 0
