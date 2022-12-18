import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class VPNWindow():
        def __init__(self,VPNObj) -> None:
            self.VPN = VPNObj
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

        def _GetPath(self):
            self.VPN.ovpnDir = filedialog.askdirectory()
            self.getPath.configure(text="已選擇檔案")
            self.OKBTN.configure(state="normal")
        
        def _OK(self):
            if (self.UserName.get() != "") and (self.Password.get() != "") and self._fileCK():
                with open("login.conf","w",encoding="utf-8") as f :
                    data = str(self.UserName.get()).strip() + "\n" + str(self.Password.get()).strip()
                    f.write(data)
                self.root.destroy() # 刪除視窗
                self.root.quit() # 離開本物件回到上層物件
            else:
                messagebox.showerror(title="檢查您的輸入",message="請檢查您的帳號或密碼是否有輸入完成,或是檔案路徑是否有效(必須包含副檔名為.ovpn的檔案類型)!!!")

        def _fileCK(self):
            self.VPN.fileList = os.listdir(self.VPN.ovpnDir)
            for file in self.VPN.fileList:
                if ".ovpn" in file:
                    return True
            return False