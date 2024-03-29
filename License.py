# 授權驗證
import datetime
import ntplib


class License():
    def __init__(self,mainWindowsObj):
        self.windows = mainWindowsObj

    def TimeLicense(self,day):
        hosts = ['time.asia.apple.com','time.google.com','time.stdtime.gov.tw','time.windows.com']
        t = ntplib.NTPClient()
        for host in hosts:
            try:
                r = t.request(host)
                break
            except:
                pass
        t = r.tx_time
        date1,time1 = str(datetime.datetime.fromtimestamp(t))[:22].split(' ')
        day1 = datetime.datetime.strptime(date1,'%Y-%m-%d')
        day2 = datetime.datetime.strptime(day,'%Y-%m-%d')
        days = day2 - day1
        if(days.days <= 0):
            self.windows.LicenseErr()
        else:
            self.windows.LicenseInfo(days.days)