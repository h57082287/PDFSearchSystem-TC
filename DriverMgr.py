import requests
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile
import os

class ChromeDriverManager:
    def __init__(self, install_path:str = ".\\") -> str:
        self.chromeVer = None
        self.debug = False
        self.path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        self.URL = None
        self.install_path = install_path + "\\chromedriver.exe"
        self.platform = sys.platform 
    
    def _getChromeVersion(self) -> bool:
        state = False
        try:
            version = os.popen("reg query HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon /v version").read()
            self.chromeVer = version.split("   ")[3].strip()
        except:
            print("Get Chrome Version Error !!!! Please set Chrome installtion Path !!!")
        return state

    def _setChromePath(self, path:str):
        self.path = path
    
    def _getDownloadURL(self) -> bool:
        state = False
        r = requests.get("https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json")
        if r.status_code == 200:
            versions = r.json()['versions']
            for version in versions:
                keyword = str(self.chromeVer).split(".")[0] + "." + str(self.chromeVer).split(".")[1] + "." + str(self.chromeVer).split(".")[2]
                if self.debug:
                    print(keyword)
                if keyword in version['version']:
                    Items = version['downloads']
                    break
            for item in Items['chromedriver']:
                if item['platform'] == self.platform:
                    self.URL = item['url']
                    state = True
                    break
        return state

    def _DownloadDriver(self):
        state = False
        r = requests.get(self.URL)
        if r.status_code == 200:
            with open(".\\Temp.zip", "wb") as f:
                f.write(r.content)
                state = True
        return state
    
    def _unZipFile(self) -> bool:
        state = False
        with ZipFile(".\\Temp.zip", "r") as zipfile:
            zipfile.extractall("./")
            os.popen("move .\\chromedriver-" + self.platform + "\\chromedriver.exe .\\")
            state = True
        return state
    
    def _DelTempFile(self):
        os.popen("rmdir /q /s chromedriver-" + self.platform)
        os.popen("del /q .\\Temp.zip")
    
    def install(self) -> str:
        self._getChromeVersion()
        self._getDownloadURL()
        self._DownloadDriver()
        self._unZipFile()
        self._DelTempFile()
        return self.install_path


# mgr = ChromeDriverManager()
# mgr.install()