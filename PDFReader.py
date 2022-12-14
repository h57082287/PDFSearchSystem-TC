import pdfplumber
import time

class PDFReader():
    def __init__(self,mainWindowsObj, filepath:str):
        self.window = mainWindowsObj
        self.FilePath = filepath
        self.PersionalData = {}

    def GetData(self,num) -> bool:
        self.window.setStatusText(content="~PDF資料解析中~",x=0.3,y=0.7,size=24)
        try:
            self.Datas = []
            pdf =  pdfplumber.open(self.FilePath)
            first_page = pdf.pages[num]
            tables = first_page.extract_table()
            for i in range(4):
                tables.pop(0)
            
            LineNum = 0
            non_data = 0
            for table  in tables:
                if LineNum%3 == 0:
                    self.PersionalData['Name'] = table[0].replace('\n','').replace(' ','')
                    self.PersionalData['Sex'] = table[1].replace('\n','').replace(' ','')
                    self.PersionalData['Born'] = table[2].replace('\n','').replace(' ','')
                    for t in table:
                        try:
                            if (t != None) and (t != "") and (len(t) == 10 and int(t[1:]) and t[0] != "▲"):
                                self.PersionalData['ID'] = t.replace('\n','').replace(' ','')
                        except:
                            pass
                        # try:
                        #     if (t != None) and (t != "") and (len(t.split('/')) == 3):
                        #         self.PersionalData['Born'] = t.replace('\n','').replace(' ','')
                        # except:
                        #     pass
                elif LineNum%3 == 1:
                    self.PersionalData['Address'] = table[1].replace('\n','').replace(' ','')
                elif LineNum%3 == 2:
                    try:
                        if (self.PersionalData['Name'] == '' or self.PersionalData['Born'] == '' or self.PersionalData['ID'] == '' or not self._CheckID(self.PersionalData['ID'])):
                            non_data += 1
                            self.PersionalData.clear()
                            LineNum+=1
                            continue 
                    except:
                        non_data += 1
                        self.PersionalData.clear()
                        LineNum+=1
                        continue
                    self.Datas.append(self.PersionalData.copy())
                    self.PersionalData.clear()
                LineNum += 1
            
            self.window.setStatusText(content="~本頁含有"+ str(non_data) +"筆資料不齊或內容異常，已跳過異常及不齊全資料~",x=0.1,y=0.8,size=12)
            time.sleep(5)
            return True,self.Datas,len(self.Datas)
        except:
            self.window.setStatusText(content="~PDF解析異常，請檢查您的檔案，本系統將跳過此頁繼續執行~",x=0.1,y=0.8,size=12)
            time.sleep(5)
            return False,self.Datas,-1

    # 檢查身分證字號
    def _CheckID(self,ID):
        if (len(ID) != 10):
            return False
        else:
            num = ord(ID[0])
            if num == 73 :
                ID = ID.replace(ID[0],'34')
            elif num == 79:
                ID = ID.replace(ID[0],'35')
            elif num <= 72:
                ID = ID.replace(ID[0],str(num-55))
            elif num >=74 and num <= 78 :
                ID = ID.replace(ID[0],str(num-56))
            elif num >= 80 :
                ID = ID.replace(ID[0],str(num-57))
            W = [1,9,8,7,6,5,4,3,2,1,1]
            reslut = 0 
            for i in range(len(ID)) :
                reslut = reslut + int(ID[i]) * W[i]
            if reslut % 10 == 0 :
                return True
            else:
                return False
