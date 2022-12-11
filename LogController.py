import datetime

class Log():
    def __init__(self) -> None:
        self.idx = 1

    def write(self,name,id,hospital,born,page,num):
        with open((hospital+"_"+str(datetime.datetime.strftime("%Y-%m-%d"))),"a+",encoding="big5") as f :
            data = str(self.idx) + "          " + name + "          " + born + "          " + id + "          頁" + page + "          筆" + num + "\n"
            f.write(data)
            self.idx += 1