import os
from zipfile import ZipFile
import requests
from datetime import datetime
import sys
import tkinter.messagebox

class UpdateController:
    def __init__(self, name: str=None, current_Version: str=None, sw_id: str=None) -> None:
        self.link = "https://updatequery-t5ljyiai2q-uc.a.run.app"
        self.remote_version = None
        argments = sys.argv
        argments.pop(0)

        if(len(argments) % 2 != 0):
            return
        if len(argments) > 1:
            try:
                # 讀入參數
                for  idx in range(0, len(argments), 2):
                    self.__settingFunction(argments[idx], argments[idx+1])
                # 執行更新
                self.__Update()
            except:
                tkinter.messagebox.showerror(title="更新失敗", message= "您的更新無法順利執行 !!!")
                print("更新失敗")
        else:
            self.SW_name= name
            self.current_Version= current_Version
            self.mode= "release"
            self.sw_id = sw_id
            self.key = "wfhiw545W152nhcue264"
        
        self.data = {
            "timestamp": str(datetime.now()),
            "sw_name": self.SW_name,
            "sw_ver": self.current_Version,
            "mode": self.mode,
            "sw_id": self.sw_id,
            "token": self.key
        }
    
    def __settingFunction(self, arg, data):
        if arg == "-cv":
            self.current_Version = data
        elif arg == "-m":
            self.mode = data
        elif arg == "-n":
            self.SW_name = data
        elif arg == "-d":
            self.sw_id = data
        elif arg == "-k":
            self.key = data
        elif arg == "-rv":
            self.remote_version = data
        elif arg == "-i":
            print("UpdateControl V1.0.0 for [PDFSearchSystem]")
        elif arg == "-h":
            print("-cv [data]  current software version")
            print("-m [data]  support mode release/debug")
            print("-i [data]  About UpdateController Version")
            print("-n [data]  software name")
            print("-d [data]  software ID")
            print("-k [data]  Verify token")
            print("-rv [data]  Update to this version")

    def __Update(self):
        if self.__ProcessCK(self.SW_name):
            if self.__KillProcess(self.SW_name):
                print("已成功" + self.SW_name + "中斷程式")
            else:
                print(self.SW_name + "中斷失敗，請手動中斷再執行更新!!!")
                return
        
        downloadLink, name, fileType = self.__sendRequest()
        if downloadLink != None:
            if self.__Download(downloadLink, name, fileType):
                print("下載完成")
            else:
                print("下載失敗，請稍後再嘗試")
        
        if fileType == "zip":
            if self.__unZipFile(name):
                print("解壓縮成功")
            else:
                print("解壓縮失敗請自行解壓縮檔案")
        else:
            print("更新完成")
            os.system(name + ".exe")
        
    def __sendRequest(self):
        try:
            r = requests.post(self.link, data=self.data)
            if r.status_code == 200:
                data = r.json()
                return data["downloadLink"], data["new_name"], data["fileType"]
            return None, None, None
        except:
            print("請求無效")
    
    def __Download(self, link, name, fileType) -> bool:
        r = requests.get(link)
        try:
            with open(name + "." + fileType, "wb") as f:
                f.write(r.content)
                return True
        except:
            return False
    
    # 檢查軟體執行狀態
    def __ProcessCK(self,ProcessName) -> bool:
        cmd = "tasklist"
        Tasks = os.popen(cmd).read()
        for Process in Tasks.split("\n"):
            if ProcessName in Process:
                print(ProcessName)
                return True
        return False
    
    # 強制中斷指定程式
    def __KillProcess(self, Processname: str) -> bool:
        cmd = "taskkill /IM " + Processname + " /F"
        res = os.popen(cmd).read()
        if "成功" in res:
            return True
        return False
    
    # 檔案解壓縮
    def __unZipFile(self, name) -> bool:
        state = False
        with ZipFile(".\\" + name + ".zip", "r") as zipfile:
            zipfile.extractall("./")
            for file in os.listdir(".\\" + name):
                os.popen("move " + file + " .\\")
            self.__DelTempFile(name)
            state = True
        return state

    # 刪除暫存檔案
    def __DelTempFile(self, name):
        os.popen("rmdir /q /s " + name)
        os.popen("del /q .\\" + name + ".zip")

if __name__ == "__main__":
    try:
        if tkinter.messagebox.askyesno(title="更新助手", message="是否開始檢查更新，一旦按下yes將關閉正在執行的所有程式 !!!"):
            UpdateController()
        else:
            os._exit(0)
    except:
        pass