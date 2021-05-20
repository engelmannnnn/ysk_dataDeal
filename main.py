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

    def dataDeal(self, startCell, endCell, outCell, method, resultsNum=3):
        startChr, startNUm = startCell[0], startCell[1:]
        endChr, endNUm = endCell[0], endCell[1:]
        resultDatas = []

        self.app.print_text("运算开始：")
        self.app.print_text("--------------------------")
        # 记录标准差大于0.05的数据集合
        discard_nums = []
        for column in range(ord(startChr), ord(endChr)+1):
            # 记录每一列数据
            nums = []
            for row in range(int(startNUm), int(endNUm)+1):
                try:
                    res = self.xls.readXls(chr(column)+str(row))
                    if isinstance(res, float) or isinstance(res, int):
                        nums.append(res)
                except Exception:
                    self.app.print_text("Error: 数据导入失败.. 失败单元格: {}".format(chr(column)+str(row)))

            # ("自定义样本标准差: 3","众数: 3","平均数" ,"样本标准差","样本方差", "总体标准差", "总体方差") mode average stdev variance pstdev pvariance
            if method == "mode":
                resData = dataDeal.modeNums(nums = nums, resultsNum=resultsNum)
                resultDatas.append(resData)
                self.app.print_text(str(resData))
            elif method == "customStdev":
                status, stdev_res, resData = dataDeal.customStdev(nums=nums, resultsNum=resultsNum)
                resultDatas.append(resData)
                if status == False:
                    if stdev_res <= 0.05:
                        self.app.print_text("注意：{}输入数据个数小于设置的输出数据个数,该组数据最小标准差小于等于0.05为：{}，已全部输出在表格中..".format(nums, round(stdev_res,3)))
                    elif stdev_res == 400:
                        self.app.print_text("注意：{}输入数据个数小于设置的输出数据个数,且该组数据最小标准差大于0.05为：{}， 该组数据已舍弃..".format(nums,  round(stdev_res,3)))
                        discard_nums.append(nums)
                    else:
                        self.app.print_text("注意：{}该组数据最小标准差大于0.05为：{},该组数据已舍弃..".format(nums,  round(stdev_res,3)))
                        discard_nums.append(nums)
                else:
                    self.app.print_text(str(resData))
            elif method == "average":
                resData = dataDeal.avgNum(nums=nums)
                resultDatas.append(resData)
                self.app.print_text(str(resData))
            elif method == "stdev":
                resData = dataDeal.stdevNum(nums=nums)
                resultDatas.append(resData)
                self.app.print_text(str(resData))
            elif method == "variance":
                resData = dataDeal.varianceNum(nums=nums)
                resultDatas.append(resData)
                self.app.print_text(str(resData))
            elif method == "pstdev":
                resData = dataDeal.pstdevNum(nums=nums)
                resultDatas.append(resData)
                self.app.print_text(str(resData))
            elif method == "pvariance":
                resData = dataDeal.pvarianceNum(nums=nums)
                resultDatas.append(resData)
                self.app.print_text(str(resData))

        self.app.print_text("--------------------------")
        self.app.print_text("运算结束")
        if method == "customStdev":
            self.app.print_text("因标准差大于0.05，丢弃的数据: \n{}".format(discard_nums))

        # 将运算结果写入表格
        outChr, outNum = outCell[0], outCell[1:]
        for column in range(len(resultDatas)):
            for row in range(len(resultDatas[column])):
                if ord(outChr)+column >= 65 and ord(outChr)+column <= 90:
                    cell = chr(ord(outChr)+column) + str(int(outNum)+row)
                elif ord(outChr)+column >90 and ord(outChr)+column <= (90+26):
                    outChr2 = chr(ord(outChr)+column - 90 + 64)
                    cell = "A"+ outChr2 + str(int(outNum)+row)
                    print(ord(outChr)+column - 90 + 64)
                    print(cell)
                else:
                    cell = ""
                    self.app.print_text("Error: Excel输出超过最大列数")
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
            method = self.app.calMethodInput.get()
            # ("自定义样本标准差: 3","众数: 3","平均数" ,"样本标准差","样本方差", "总体标准差", "总体方差") customStdev mode average stdev variance pstdev pvariance
            if method.find("众数") != -1:
                resultsNum = int(method.split(":")[-1])
                if resultsNum == 0:
                    self.app.print_text("Error: 取值不能为0")
                else:
                    self.app.print_text("众数取值: {}".format(resultsNum))
                    self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="mode", resultsNum=resultsNum)
            elif method.find("自定义样本标准差") != -1:
                resultsNum = int(method.split(":")[-1])
                if resultsNum == 0:
                    self.app.print_text("Error: 取值不能为0")
                else:
                    self.app.print_text("样本标准差取值: {}".format(resultsNum))
                    self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="customStdev", resultsNum=resultsNum)
            elif method == "平均数":
                self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="average")
            elif method == "样本标准差":
                self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="stdev")
            elif method == "样本方差":
                self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="variance")
            elif method == "总体标准差":
                self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="pstdev")
            elif method == "总体方差":
                self.dataDeal(startCell=startCell, endCell=endCell, outCell=outCell, method="pvariance")







    def MainApp(self):
        # 创建一个MainUI对象
        # self.app = MainUI()
        # 设置窗口标题
        self.app.master.title('科学计算器')
        # 设置窗体大小
        self.app.master.geometry('777x555+500+250')
        self.app.master.resizable(False, False)


        # 补充配置,补充gui中button_run按钮的command参数
        self.app.button_run.config(command=self.run)
        # 主循环开始
        self.app.mainloop()

if __name__ == "__main__":
    tst = mainFunc()
    tst.MainApp()
    # filePath = "E:/PyProject/ysk_dataDeal/result/毒性增殖试验.xlsx"
    # tst.openXls(filePath, "Sheet1")
    # tst.dataDeal("B169","H173","J169")

