from openpyxl import load_workbook
import traceback
import time

class excelControl():

    def openXls(self, filePath, sheetName):
        try:
            self.fileDir = ""
            for dir in filePath.split("/")[:-1]:
                self.fileDir += dir+"/"
            self.fileName = filePath.split("/")[-1]
            self.workbook = load_workbook(filePath)
            self.worksheet = self.workbook[sheetName]

        except Exception:
            print("文件打开错误, 文件未找到 或 文件内表格(sheet)名称有误! 请注意大小写. \n文件路径: {} 表格名称: {}".format(filePath,sheetName))

    # def readXls(self, column: int, row: int):
    #     col = chr(column + 64)
    #     cell = str(col + str(row))
    #     return self.worksheet[cell].value

    def readXls(self, cell):
        return self.worksheet[cell].value


    def writeXls(self,cell, data):
        try:
            self.worksheet[cell] = data
        except Exception:
            print("数据写入错误, 未写入数据: {}".format(data))
            print(traceback.format_exc())

    def saveXls(self):
        now = time.strftime("%Y%m%d-%H%M%S")
        fileName = now+self.fileName
        filePath = self.fileDir+fileName
        self.workbook.save(filePath)
        print("文件保存路径: {}".format(filePath))

if __name__ == "__main__":
    tst = excelControl()
    filePath = "E:/PyProject/ysk_dataDeal/result/毒性增殖试验.xlsx"
    path2 = "E:/PyProject/DataSpider/result/Marketplace_salesforce.xlsx"
    tst.openXls(filePath,"Sheet1")
    tst.saveXls()