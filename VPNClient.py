import os 
import time
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup
import tkinter as tk

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
        self.fileList = []
        self.index = 0

        if (self.window.checkVal_AUVPNM.get()):
            # 設定視窗
            self.root = tk.Tk()
            self.root.title("VPN設定視窗")
            self.root.geometry("300x200")
            self.root.resizable(width=False,height=False)

            # 設定文字
            self.Text1 = tk.Label(self.root,text="OpenVPN帳號 : " ,font=("標楷體", 12))
            self.Text1.place(relx=0.05,rely=0.2)
            self.UserName = tk.Entry(self.root)
            self.UserName.place(relx=0.45,rely=0.2)

            # 設定文字
            self.Text2 = tk.Label(self.root,text="OpenVPN密碼 : " ,font=("標楷體", 12))
            self.Text2.place(relx=0.05,rely=0.35)
            self.Password = tk.Entry(self.root,show='*')
            self.Password.place(relx=0.45,rely=0.35)

            # 建立按鈕
            self.getPath = tk.Button(self.root,text="選擇.ovpn資料夾路徑",width=20,height=2,font=("標楷體",12),command=self._GetPath)
            self.getPath.place(relx=0.25 , rely=0.5)

            # 建立按鈕
            self.OKBTN = tk.Button(self.root,text="確認",width=7,height=1,font=("標楷體",12),command=self._OK,state="disable")
            self.OKBTN.place(relx=0.37 , rely=0.78)

            self.root.mainloop()
        
    def _InstallationCkeck(self):
        cmd = 'start openvpn --config "C:\\Users\\h5708\\Downloads\\Surfshark_Config\\89.187.178.94_tcp.ovpn" --auth-user-pass login.conf'
        os.system(cmd)
        time.sleep(5)
        if self._CKVPNStatus:
            self._stopVPN()
        else:
            pass

    def _GetPath(self):
        self.ovpnDir = filedialog.askdirectory()
        self.getPath.configure(text="已選擇檔案")
        self.OKBTN.configure(state="normal")
    
    def _OK(self):
        if (self.UserName.get() != "") and (self.Password.get() != "") and self._fileCK():
            with open("login.conf","w",encoding="utf-8") as f :
                data = str(self.UserName.get()).strip() + "\n" + str(self.Password.get()).strip()
                f.write(data)
            self.root.destroy()
        else:
            messagebox.showerror(title="檢查您的輸入",message="請檢查您的帳號或密碼是否有輸入完成,或是檔案路徑是否有效(必須包含副檔名為.ovpn的檔案類型)!!!")

    def _startVPN(self):
        if (self.index != len(self.fileList)) and self.window.checkVal_AUVPNM.get():
            cmd = 'start openvpn --config "' + self.fileList[self.index] + '" --auth-user-pass login.conf'
            os.system(cmd)
            self.index += 1
            self._memberControl()
        else:
            self.window.RunStatus = False
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
    
    def _fileCK(self):
        self.fileList = os.listdir(self.ovpnDir)
        for file in self.fileList:
            if ".ovpn" in file:
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