from common.excelControl import excelControl
import traceback
from dataDeal import dataDeal
from cal_gui import MainUI
class mainFunc():
    def __init__(self):
        self.xls = excelControl()
        self.app = MainUI()

    def openXls(self,filePath, sheetName):
        self.xls.openXls(filePath, sheetName)

    def dataDeal(self, startCell, endCell, outCell):
        startChr, startNUm = startCell[0], startCell[1:]
        endChr, endNUm = endCell[0], endCell[1:]
        resultDatas = []

        self.app.print_text("--------------------------\n运算结果:")
        for column in range(ord(startChr), ord(endChr)+1):
            nums = []
            for row in range(int(startNUm), int(endNUm)+1):
                try:
                    res = self.xls.readXls(chr(column)+str(row))
                    if isinstance(res, float) or isinstance(res, int):
                        nums.append(res)
                except Exception:
                    self.app.print_text("Error: 数据导入失败.. 失败单元格: {}".format(chr(column)+str(row)))
            resData = dataDeal.modeNums(nums = nums, resultsNum=3)
            resultDatas.append(resData)
            self.app.print_text(str(resData))
        self.app.print_text("--------------------------")
        self.app.print_text("运算结束\n")

        outChr, outNum = outCell[0], outCell[1:]
        for column in range(len(resultDatas)):
            for row in range(len(resultDatas[column])):
                cell = chr(ord(outChr)+column) + str(int(outNum)+row)
                self.xls.writeXls(cell=cell, data=resultDatas[column][row])

        # 存储表格, 可自定义是否另存为
        if self.app.choose.get():
            state, filePath = self.xls.saveXls(type="new")
            if state:
                self.app.print_text("文件另存为成功, 文件地址:")
                self.app.print_text(filePath)
            else:
                self.app.print_text("Error: 文件另存为失败, 应存的文件地址:")
                self.app.print_text(filePath)
        else:
            state, filePath = self.xls.saveXls(type="origin")
            if state:
                self.app.print_text("保存原文件成功, 文件地址:")
                self.app.print_text(filePath)
            else:
                self.app.print_text("Error: 保存原文件失败, 请检查原文件是否已关闭, 未关闭时无法保存..")
                self.app.print_text("文件已另存为, 文件地址: ")
                self.app.print_text(filePath)

    def run(self):
        # 获取页面参数
        startChr = self.app.dataBeginInput.get()
        startNum = self.app.dataBeginNumInput.get()
        endChr = self.app.dataEndInput.get()
        endNum = self.app.dataEndNumInput.get()
        outChr = self.app.dataEndOutput.get()
        outNum = self.app.dataEndNumOutput.get()

        startCell = startChr + startNum
        endCell = endChr + endNum
        outCell = outChr + outNum

        filePath = self.app.path.get()
        sheetName = self.app.sheetNameInput.get()
        if startNum == '' or endNum == '' or outNum == '':
            self.app.print_text("Error: 请输入正确的单元格..")
        elif self.app.path.get() == '':
            self.app.print_text("Error: 请选择正确的文件路径..")

        else:
            # 操作文件流
            try:
                self.openXls(filePath, sheetName)
            except Exception:
                self.app.print_text("文件解析失败..")
                self.app.print_text(traceback.format_exc())

            if self.app.calMethodInput.get() == "取众数":
                self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell)



    def MainApp(self):
        # 创建一个MainUI对象
        # self.app = MainUI()
        # 设置窗口标题
        self.app.master.title('科学计算器')
        # 设置窗体大小
        self.app.master.geometry('777x555+500+250')
        self.app.master.resizable(False, False)


        # 补充配置,补充gui中button_run按钮的command参数
        if self.app.calMethodInput.get() == "取众数":
            self.app.button_run.config(command=self.run)
        # 主循环开始
        self.app.mainloop()

if __name__ == "__main__":
    tst = mainFunc()
    tst.MainApp()
    # filePath = "E:/PyProject/ysk_dataDeal/result/毒性增殖试验.xlsx"
    # tst.openXls(filePath, "Sheet1")
    # tst.dataDeal("B169","H173","J169")

