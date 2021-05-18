from common.excelControl import excelControl
import traceback
from dataDeal import dataDeal


class mainFunc():
    def __init__(self):
        self.xls = excelControl()

    def openXls(self,filePath, sheetName):
        self.xls.openXls(filePath, sheetName)

    def dataDeal(self, startCell, endCell, outCell):
        startChr, startNUm = startCell[0], startCell[1:]
        endChr, endNUm = endCell[0], endCell[1:]
        resultDatas = []

        for column in range(ord(startChr), ord(endChr)+1):
            nums = []
            for row in range(int(startNUm), int(endNUm)+1):
                try:
                    res = self.xls.readXls(chr(column)+str(row))
                    if isinstance(res, float) or isinstance(res, int):
                        nums.append(res)
                except Exception:
                    print("数据导入失败.. 失败单元格: {}".format(chr(column)+str(row)))
            resData = dataDeal.modeNums(nums = nums, resultsNum=3)
            resultDatas.append(resData)
            print(resultDatas)

        outChr, outNum = outCell[0], outCell[1:]
        for column in range(len(resultDatas)):
            for row in range(len(resultDatas[column])):
                cell = chr(ord(outChr)+column) + str(int(outNum)+row)
                self.xls.writeXls(cell=cell, data=resultDatas[column][row])
            print("----")
        self.xls.saveXls()

if __name__ == "__main__":
    tst = mainFunc()
    filePath = "E:/PyProject/ysk_dataDeal/result/毒性增殖试验.xlsx"
    tst.openXls(filePath, "Sheet1")
    tst.dataDeal("B169","H173","J169")

