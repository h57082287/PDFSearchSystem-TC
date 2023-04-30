import ctypes
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import tkinter
from tkinter import messagebox
from License import License
from auh import AUH
from ckccgh import CKCCGH
from ptccgh import PTCCGH
from tzuchi import TZUCHI
from ujah import UJAH
from jjah import JJAH
from cmuh import CMUH
from cmuhh import CMUHH
from fyh import FYH
from rg import RG
from wlshosp import WLSHOSP
from mohw import MOHW
from lshosp import LSHOSP
from everan import EVERAN
from chingchyuan import CHINGCHYUAN
from tafghzb import TAFGHZB
from tafgh import TAFGH
from csh import CSH
from vghtc import VGHTC
from ktghs import KTGHS
from ktghd import KTGHD
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
# from VPNClient import VPN

# ========================================================
# 注意 : 每次添加新的醫院時須同時調整所有的URLList
# ========================================================
# 建立視窗程式物件
class MainWindows():
    def __init__(self):
        
        # 檢查是否多開
        # if self._MultiProcessCK("PDFSearchSystem-TC.exe") :
        #     self.MultiProcessErr()

        # 取得管理員權限
        self.getAdmin()
        
        # 變數設定
        self.mode = 0
        self.modeName = '全部'
        self.URLList = {
            "童綜合醫院" : None,
            "亞洲大學附設醫院" : None,
            "澄清醫院中港分院" : None,
            "澄清醫院" : None,
            "台中仁愛醫院" : None,
            "大里仁愛醫院" : None,
            "中國醫學大學豐原分院" : None,
            "中國醫學大學" : None,
            "林新醫院" : None,
            "部立台中醫院" : None,
            "慈濟醫院台中分院" : None,
            "豐原醫院" : None,
            "太平長安醫院" : None,
            "清泉醫院" : None,
            "烏日林新醫院" : None,
            "國軍醫院-中清" : None,
            "國軍醫院-台中" : None,
            "中山醫" : None,
            "台中榮總" : None,
            # "光田醫院：沙鹿總院" : None,
            # "光田醫院：大甲院區" : None,
        }
        self.RunStatus = True
        self.VClient = None

        # 設定視窗
        self.window = tk.Tk()
        self.window.title("PDF比對搜尋系統-台中地區")
        self.window.geometry("600x600")
        self.window.resizable(width=False,height=False)
        
        #-----------------------------------------------------------------------------------------標題
        # 設定標題文字
        self.tittle = tk.Label(self.window,text="PDF比對搜尋系統-台中地區",font = ("標楷體" , 28))
        self.tittle.pack(side=tk.TOP)

        #-----------------------------------------------------------------------------------------標題
        # 建立tab標籤控制器
        self.TabControl = ttk.Notebook(self.window)
        
        # 建立標籤-mainTab
        self.MainTab = tk.Frame(self.TabControl)
        self.TabControl.add(self.MainTab,text="主介面")
        self.TabControl.pack(expand=1, fill="both")

        # 建立標籤-setup
        self.SetupTab = tk.Frame(self.TabControl)
        self.TabControl.add(self.SetupTab,text="設定")
        self.TabControl.pack(expand=1, fill="both")
        #-----------------------------------------------------------------------------------------PDF選擇
        # 設定文字
        self.Text1 = tk.Label(self.MainTab,text="請選擇PDF檔案 : " ,font=("標楷體", 18))
        self.Text1.place(relx=0.2,rely=0.10,anchor=tk.CENTER)

        # 設定文字
        self.fileText = tk.Label(self.MainTab,text="-" , font=("Times New Roman" , 10))
        self.fileText.place(relx=0.4,rely=0.10,anchor=tk.CENTER)

        # 設定按鈕
        self.getFile = tk.Button(self.MainTab,text="選擇檔案",width=8,height=2,font=("標楷體",12),command=self.GetFile)
        self.getFile.place(relx=0.8 , rely=0.06)

        #-----------------------------------------------------------------------------------------PDF設定
        # 設定文字
        self.Text4 = tk.Label(self.MainTab,text="從第" , font=("標楷體" , 14))
        self.Text4.place(relx=0.05,rely=0.21,anchor=tk.CENTER)
        
        # 輸入啟始頁數
        self.beginPage = tk.Entry(self.MainTab,width=5)
        self.beginPage.place(relx=0.1,rely=0.19)
        self.beginPage.insert(0,'1')
        self.beginPage.configure(state='disable')

        # 設定文字
        self.Text5 = tk.Label(self.MainTab,text="頁、第" , font=("標楷體" , 14))
        self.Text5.place(relx=0.22,rely=0.21,anchor=tk.CENTER)

        # 輸入啟始項數
        self.beginNum = tk.Entry(self.MainTab,width=3)
        self.beginNum.place(relx=0.28,rely=0.19)
        self.beginNum.insert(0,'1')
        self.beginNum.configure(state='disable')

        # 設定文字
        self.Text6 = tk.Label(self.MainTab,text="項開始到第" , font=("標楷體" , 14))
        self.Text6.place(relx=0.42,rely=0.21,anchor=tk.CENTER)

        # 輸入結束頁數
        self.endPage = tk.Entry(self.MainTab,width=5)
        self.endPage.place(relx=0.5,rely=0.19)
        self.endPage.configure(state='normal')

        # 設定文字
        self.Text7 = tk.Label(self.MainTab,text="頁、第" , font=("標楷體" , 14))
        self.Text7.place(relx=0.62,rely=0.21,anchor=tk.CENTER)

        # 輸入結束項數
        self.endNum = tk.Entry(self.MainTab,width=5)
        self.endNum.place(relx=0.69,rely=0.19)
        self.endNum.configure(state='disable')

        # 設定文字
        self.Text8 = tk.Label(self.MainTab,text="項結束" , font=("標楷體" , 14))
        self.Text8.place(relx=0.80,rely=0.21,anchor=tk.CENTER)

        # 設定按鈕
        self.switchMode = tk.Button(self.MainTab,text="指定",width=6,height=1,font=("標楷體",14),command=self.changeMode)
        self.switchMode.place(relx=0.87 , rely=0.18)

        #-----------------------------------------------------------------------------------------匯出選擇
        # 設定文字
        self.Text2 = tk.Label(self.MainTab,text="請選擇匯出路徑 : " ,font=("標楷體", 18))
        self.Text2.place(relx=0.19,rely=0.32,anchor=tk.CENTER)

        # 設定文字
        self.pathText = tk.Label(self.MainTab,text="-" , font=("Times New Roman" , 10))
        self.pathText.place(relx=0.4,rely=0.32,anchor=tk.CENTER)

        # 設定按鈕
        self.getPath = tk.Button(self.MainTab,text="選擇路徑",width=8,height=2,font=("標楷體",12),command=self.GetPath)
        self.getPath.place(relx=0.8 , rely=0.28)

        #-----------------------------------------------------------------------------------------網站選擇
        # 獲取列表
        # MainWindows.GetWebList()

        # 設定文字
        self.Text3 = tk.Label(self.MainTab,text="請選擇比對網站 : " , font=("標楷體" , 18))
        self.Text3.place(relx=0.19,rely=0.43,anchor=tk.CENTER)

        # 設定下拉是選單
        self.webList = ttk.Combobox(self.MainTab,values=list(self.URLList.keys()),font=("標楷體",14),width=25,state="readonly")
        self.webList.current(0)
        self.webList.place(relx= 0.39 , rely=0.41)

        #-----------------------------------------------------------------------------------------啟動按鈕
        self.startButton = tk.Button(self.MainTab,text="開始比對",font=("標楷體",18) , width= 8 , height=1,command=self.Start)
        self.startButton.place(relx= 0.3, rely=0.50)

        #-----------------------------------------------------------------------------------------停止按鈕
        self.StopButton = tk.Button(self.MainTab,text="停止比對",font=("標楷體",18) , width= 8 , height=1,command=self.Stop)
        self.StopButton.place(relx= 0.48, rely=0.50)
        self.StopButton.config(state="disable")

        #-----------------------------------------------------------------------------------------狀態文字
        self.statusText = tk.Label(self.MainTab,text="~就緒~",font=("標楷體",24))
        self.statusText.place(relx=0.4 , rely= 0.7 )

        #-----------------------------------------------------------------------------------------關閉視窗
        # 視窗關閉
        self.window.protocol("WM_DELETE_WINDOW",self.onClose)
        
        # 檢查授權
        l = License(self)
        l.TimeLicense("2023-05-14")
        
        #-----------------------------------------------------------------------------------------啟動視窗
        # 建立設定頁面內容

        # 初始化狀態
        self.checkVal_AND = tk.BooleanVar()
        self.checkVal_AND.set(False)

        self.checkVal_AUDM = tk.BooleanVar()
        self.checkVal_AUDM.set(False)

        self.checkVal_AULL = tk.BooleanVar()
        self.checkVal_AULL.set(True)

        self.checkVal_AUVPNM = tk.BooleanVar()
        self.checkVal_AUVPNM.set(False)

        # 建立勾選
        self.allowNetworkDebug = tk.Checkbutton(self.SetupTab,text="允許使用網路偵錯功能(開發者可遠端除錯)",var=self.checkVal_AND,state="disable")
        self.allowNetworkDebug.place(relx=0.25,rely=0.08)

        self.allowUseDebugMode = tk.Checkbutton(self.SetupTab,text="允許使用測試版軟體(可能會有穩定性問題)",var=self.checkVal_AUDM,state="disable")
        self.allowUseDebugMode.place(relx=0.25,rely=0.18)

        self.allowUseLocalLog = tk.Checkbutton(self.SetupTab,text="產生本地端記錄檔(與執行檔相同資料夾)",var=self.checkVal_AULL,state="disabled")
        self.allowUseLocalLog.place(relx=0.25,rely=0.28)

        self.allowUseVPNMode = tk.Checkbutton(self.SetupTab,text="允許使用VPN繞過IP封鎖機制\n(目前僅提供中國學大學系列與林新醫院，若有其他醫院也需要再回報)",var=self.checkVal_AUVPNM)
        self.allowUseVPNMode.place(relx=0.25,rely=0.38)

        # 建立按鈕
        self.checkUpdate = tk.Button(self.SetupTab,text="檢查更新",font=("標楷體",18) , width= 8 , height=1,command=None)
        self.checkUpdate.place(relx= 0.4, rely=0.45)
        self.checkUpdate.config(state="disable")

        # 版本說明
        self.VersionInfo = tk.LabelFrame(self.SetupTab,text="版本資訊",width=350,height=230)
        self.VersionInfo.place(relx=0.2,rely=0.55)
        self.version_text = tk.Label(self.VersionInfo,text="版本 : v3.5.3" ,font=("標楷體", 12))
        self.version_text.place(relx=0.1,rely=0.05)
        self.type_text = tk.Label(self.VersionInfo,text="類型 : 試用版" ,font=("標楷體", 12))
        self.type_text.place(relx=0.1,rely=0.16)
        self.status_text = tk.Label(self.VersionInfo,text="授權 : 未授權" ,font=("標楷體", 12))
        self.status_text.place(relx=0.1,rely=0.27)
        self.release_date = tk.Label(self.VersionInfo,text="發布日期 : 2023-04-30" ,font=("標楷體", 12))
        self.release_date.place(relx=0.1,rely=0.38)
        self.evmt = tk.Label(self.VersionInfo,text="支援環境 : Windows 10 、 Windows 11" ,font=("標楷體", 12))
        self.evmt.place(relx=0.1,rely=0.49)
        self.ChangeInfo = tk.LabelFrame(self.VersionInfo,text="變更內容",width=280,height=77)
        self.ChangeInfo.place(relx=0.1,rely=0.60)
        self.ChangeInfo_content = tk.Label(self.ChangeInfo,text="修正國軍中清重試機制",font=("標楷體", 8))
        self.ChangeInfo_content.place(relx=0.1,rely=0.01)
        # ----------------------------------------------------------------------------------------
        # 選擇預設標籤
        self.TabControl.select(self.MainTab)

        self.window.mainloop()

    # 關閉功能定義
    def onClose(self):
        # MainWindows.StopThread = True
        self.RunStatus = False
        try:
            self.endBrowser()
        except:
            pass
        os._exit(0)

    # 連結取得檔案動作
    def GetFile(self):
        self.filePath = filedialog.askopenfilename()
        self.fileText.configure(text=self.filePath)
        self.fileText.place(relx= 0.6)
    
    # 連結取得路徑動作
    def GetPath(self):
        self.outputPath  = filedialog.askdirectory()
        self.pathText.configure(text=self.outputPath)
        self.pathText.place(relx= 0.6)
    
    # 連結修改模式動作
    def changeMode(self):
        if((self.mode % 2) == 0):
            self.switchMode.configure(text="全部")
            self.beginPage.configure(state="normal")
            self.beginNum.configure(state="normal")
            self.endNum.configure(state="normal")
            self.modeName = '指定'
        else:
            self.switchMode.configure(text="指定")
            self.beginPage.configure(state="disable")
            self.beginNum.configure(state="disable")
            self.endNum.configure(state="disable")
            self.modeName = '全部'
        self.mode = self.mode + 1
    
    # 介面鎖定
    def _LockUI(self):
        self.switchMode.configure(state="disable")
        self.beginPage.configure(state="disable")
        self.beginNum.configure(state="disable")
        self.endNum.configure(state="disable")
        self.endPage.configure(state="disable")
        self.getFile.configure(state="disable")
        self.getPath.configure(state="disable")
        self.startButton.configure(state="disable")
        self.StopButton.configure(state="normal")
        self.webList.configure(state="disable")

    # 連結停止動作
    def Stop(self):
        self.setStatusText(content="停止服務中...",x=0.25,y=0.7,size=24)
        self.RunStatus = False
    
    # 介面還原
    def GUIRestart(self):
        self.RunStatus = True
        self.startButton.configure(state="normal")
        self.StopButton.configure(state="disable")
        self.switchMode.configure(state="normal",text= "指定")
        self.mode = 0

        self.beginNum.configure(state="normal")
        self.beginNum.delete(0,"end")
        self.beginNum.insert(0,'1')
        self.beginNum.configure(state="disable")

        self.beginPage.configure(state="normal")
        self.beginPage.delete(0,"end")
        self.beginPage.insert(0,'1')
        self.beginPage.configure(state="disable")

        self.endPage.configure(state="normal")
        self.endPage.delete(0,"end")
        
        self.endNum.configure(state="normal")
        self.endNum.delete(0,"end")
        self.endNum.configure(state="disable")

        self.getFile.configure(state="normal")
        self.getPath.configure(state="normal")
        self.webList.configure(state="readonly")

    def Start(self):
        self._LockUI()
        self.browser = webdriver.Chrome(ChromeDriverManager().install())
        self.browser.set_page_load_timeout(180)
        print("Init URLList")
        self.URLList = {
            "童綜合醫院" : RG(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "亞洲大學附設醫院" : AUH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "澄清醫院中港分院" : CKCCGH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "澄清醫院" : PTCCGH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "台中仁愛醫院" : UJAH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "大里仁愛醫院" : JJAH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "中國醫學大學豐原分院" : CMUH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "中國醫學大學" : CMUHH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "林新醫院" : LSHOSP(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "部立台中醫院" : MOHW(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "慈濟醫院台中分院" : TZUCHI(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "豐原醫院" : FYH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "太平長安醫院" : EVERAN(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "清泉醫院" : CHINGCHYUAN(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "烏日林新醫院" : WLSHOSP(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "國軍醫院-中清" : TAFGHZB(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "國軍醫院-台中" : TAFGH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "中山醫" : CSH(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            "台中榮總" : VGHTC(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            #api+vpn"光田醫院：沙鹿總院" : KTGHS(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
            #api+vpn"光田醫院：大甲院區" : KTGHD(self.browser,self,self.beginPage.get(),self.beginNum.get(),self.endPage.get(),self.endNum.get(),self.outputPath,self.filePath),
        }
        print("start mainProcess")
        t1 = threading.Thread(target=self.mainProcess)
        t1.start()

    # 取得PDF路徑
    def getPDFPath(self):
        return self.filePath
    
    # 取得輸出路徑路徑
    def getOutPath(self):
        return self.outputPath

    # 修改顯示文字
    def setStatusText(self,content,x,y,size):
        self.statusText.configure(text=content,font=("標楷體",size))
        self.statusText.place(relx=x,rely=y)
    
    # 建立主程序
    def mainProcess(self):
        # try:
            self.URLList[self.webList.get()].run()
        # except:
        #     messagebox.showerror("發生錯誤", "請檢查您的網路是否異常，並排除後再次執行本程式，系統將於您按下[確定]後自動關閉!!!")
        #     os._exit(0)
    
    # 結束瀏覽器
    def endBrowser(self):
        self.browser.quit()
    
    # 獲取最高權限
    def getAdmin(self):
        if self.is_admin():
           pass
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
            os._exit(0)
    
    # 檢查權限
    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    # 檢查多開狀態
    def _MultiProcessCK(self,ProcessName):
        ProcessNum = 0
        cmd = "tasklist"
        Tasks = os.popen(cmd).read()
        for Process in Tasks.split("\n"):
            if ProcessName in Process:
                print(ProcessName)
                ProcessNum += 1
        print(ProcessNum)
        if ProcessNum > 1:
            return True
        return False
    # =================================================================
    #                           彈窗定義
    # =================================================================
    def LicenseErr(self):
        tkinter.messagebox.showerror(title="授權失敗",message="注意您的有效授權已失效，請聯繫開發人員了解 !!!")
        os._exit(0)

    def LicenseInfo(self,days):
        tkinter.messagebox.showinfo(title="授權通知",message="注意，本軟體僅供測試使用，因此設有授權保護。\n您的有效授權還剩下"+ str(days) +"天!!!")
    
    def MultiProcessErr(self):
        tkinter.messagebox.showerror(title="軟體多開守門員",message="偵測到本軟體已在運行，請先將原先軟體關閉後再開啟程式 !!!")
        os._exit(0)
    
    # =================================================================
    #                           文字定義
    # =================================================================

# 啟動主程序
if __name__ == '__main__':
    mainWindows = MainWindows()