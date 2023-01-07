from datetime import datetime

class Log():
    def __init__(self) -> None:
        self.idx = 1

    def write(self,name,id,hospital,born,page,num):
        with open((hospital+"_"+str(datetime.now().strftime("%Y-%m-%d"))+".txt"),"a+",encoding="utf-8") as f :
            data = str(self.idx) + "          " + name + "          " + born + "          " + id + "          頁" + page + "          筆" + num + "\n"
            f.write(data)
            self.idx += 1