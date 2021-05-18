import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk

# 定义MainUI类表示应用/窗口，继承Frame类
class MainUI(tk.Frame):
    # Application构造函数，master为窗口的父控件
    def __init__(self, master=None):
        # 初始化Application的Frame部分
        tk.Frame.__init__(self, master)
        # 显示窗口，并使用place布局
        self.place(x=0,y=0)
        self.path = tk.StringVar()
        # 创建控件
        self.createWidgets()

    def selectPath(self):
        '''选择要处理的excel地址'''
        self.path_ = askopenfilename(title='选择文件', filetypes=[('xlsx文件', '*.xlsx'),('所有文件','*')])
        if self.path_ == "":
            self.print_text("未导入Excel！")
        else:
            self.path.set(self.path_)
            self.print_text("导入Excel成功!"+"\n"+"文件路径: "+self.path_)

    # 在gui信息栏打印输出
    def print_text(self,msg):
        self.detail_Text.insert('end', str(msg) + '\n')

    # 创建控件
    def createWidgets(self):
        font = ('微软雅黑',12,'bold')
        '''生成gui界面'''
        # 创建一个顶部标签栏
        self.top_frame = tk.Frame(self.master, bg = '#EFEEEE', width = 777, height = 40)
        self.top_frame.pack()
        self.welcome_text = tk.Label(self.top_frame, text='Excel数据处理器', bg='#EFEEEE', fg='#575757', font=font)
        self.welcome_text.place(x=20, y=8)

        # 中间空间区域
        # 创建控件集成区域 用于布局中间控件部分
        self.controlerArea = tk.Frame(self.master, width=777, height=200, bg='#F6F7F5')
        self.controlerArea.pack()

        # 数据起始单元 模块
        self.dataBegin = tk.Label(self.controlerArea, text='数据起始单元:', font=font, bg='#F6F7F5', fg='#3E3E3E')
        self.dataBegin.place(x=20, y=10)
        # V2.0版本中使用的输入框模式,2.1版本改为下拉菜单选项
        # self.testCreatorInput = tk.Entry(self.controlerArea, width=25, font=('微软雅黑', 12), bg='#F6F7F5', fg='#3E3E3E')
        # self.testCreatorInput.place(x=125, y=10)
        # V2.1版本Combobox组件,设计成下拉菜单
        self.dataBeginInput = ttk.Combobox(self.controlerArea,width=10, font=('微软雅黑', 12))
        self.dataBeginInput["value"] = ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")
        self.dataBeginInput.current(0)
        self.dataBeginInput.place(x=130, y=10)

        self.dataBeginNumInput = ttk.Combobox(self.controlerArea, width=9, font=('微软雅黑', 12))
        self.dataBeginNumInput["value"] = ("")
        self.dataBeginNumInput.place(x=250, y=10)

        # 数据截止单元 模块
        self.dataEnd = tk.Label(self.controlerArea, text='数据截止单元:', font=font, bg='#F6F7F5', fg='#3E3E3E')
        self.dataEnd.place(x=410, y=10)
        # V2.0版本中使用的输入框模式,2.1版本改为下拉菜单选项
        # self.reqIdInput = tk.Entry(self.controlerArea,width=25, font=('微软雅黑', 12),bg='#F6F7F5', fg='#3E3E3E')
        # self.reqIdInput.place(x=510, y=10)
        # V2.1版本Combobox组件,设计成下拉菜单 为了统一ui,也做了下拉框,可以在下拉框里直接输入
        self.dataEndInput = ttk.Combobox(self.controlerArea, width=10, font=('微软雅黑', 12))
        self.dataEndInput["value"] = ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")
        self.dataEndInput.current(0)
        self.dataEndInput.place(x=520, y=10)

        self.dataEndNumInput = ttk.Combobox(self.controlerArea, width=9, font=('微软雅黑', 12))
        self.dataEndNumInput["value"] = ("")
        self.dataEndNumInput.place(x=640, y=10)

        # 数据输出单元 模块
        self.dataOutput = tk.Label(self.controlerArea, text='数据输出单元:', font=font, bg='#F6F7F5', fg='#3E3E3E')
        self.dataOutput.place(x=20, y=58)
        # V2.0版本中使用的输入框模式,2.1版本改为下拉菜单选项
        # self.testTypeInput = tk.Entry(self.controlerArea, width=25, font=('微软雅黑', 12), bg='#F6F7F5', fg='#3E3E3E')
        # self.testTypeInput.place(x=125, y=58)
        # V2.1版本Combobox组件,设计成下拉菜单
        self.dataEndOutput = ttk.Combobox(self.controlerArea, width=10, font=('微软雅黑', 12))
        self.dataEndOutput["value"] = ("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z")
        self.dataEndOutput.current(0)
        self.dataEndOutput.place(x=130, y=58)

        self.dataEndNumOutput = ttk.Combobox(self.controlerArea, width=9, font=('微软雅黑', 12))
        self.dataEndNumOutput["value"] = ("")
        self.dataEndNumOutput.place(x=250, y=58)

        # 表单名称 模块
        self.sheetName = tk.Label(self.controlerArea, text='表单名称:', font=font, bg='#F6F7F5', fg='#3E3E3E')
        self.sheetName.place(x=410, y=58)
        # V2.0版本中使用的输入框模式,2.1版本改为下拉菜单选项
        # self.testLevelInput = tk.Entry(self.controlerArea, width=25, font=('微软雅黑', 12), bg='#F6F7F5', fg='#3E3E3E')
        # self.testLevelInput.place(x=510, y=58)
        # V2.1版本Combobox组件,设计成下拉菜单
        self.sheetNameInput = ttk.Combobox(self.controlerArea, width=22, font=('微软雅黑', 12))
        self.sheetNameInput["value"] = ("Sheet1")
        self.sheetNameInput.current(0)
        self.sheetNameInput.place(x=521, y=58)

        # 运算方法 模块
        self.calMethod = tk.Label(self.controlerArea, text='运算方法:', font=font, bg='#F6F7F5', fg='#3E3E3E')
        self.calMethod.place(x=20, y=105)
        # V2.0版本中使用的输入框模式,2.1版本改为下拉菜单选项
        # self.testStatusInput = tk.Entry(self.controlerArea, width=25, font=('微软雅黑', 12), bg='#F6F7F5', fg='#3E3E3E')
        # self.testStatusInput.place(x=125, y=105)
        # V2.1版本Combobox组件,设计成下拉菜单
        self.calMethodInput = ttk.Combobox(self.controlerArea, width=23, font=('微软雅黑'))
        self.calMethodInput["value"] = ("取众数","待更新")
        self.calMethodInput.current(0)
        self.calMethodInput.place(x=125, y=105)

        # 是否另存为勾选框
        self.choose = tk.IntVar()
        self.choose.set(1)
        self.saveButtonInput = tk.Checkbutton(self.controlerArea, variable=self.choose, onvalue=1, offvalue=0, text="另存为一个新的表格文件", font= ('微软雅黑',12), bg='#F6F7F5', fg='#3E3E3E')
        self.saveButtonInput.place(x=545, y=105)

        # Excel地址展示栏和地址选择按钮
        self.firstLabel = tk.Label(self.controlerArea, text='Excel地址:', font=font, bg='#F6F7F5', fg='#3E3E3E')
        self.firstLabel.place(x=20, y=150)
        self.address_show = tk.Entry(self.controlerArea,textvariable=self.path,width=62,font=('微软雅黑', 12), bg='#F6F7F5', fg='#3E3E3E')
        self.address_show.place(x=125, y=150)
        self.clickButton = tk.Button(self.controlerArea, text="浏览", font=('微软雅黑', 11, 'bold'), bg = '#778899', fg = 'white', activebackground = '#87CEFA', activeforeground = 'white', width=5, height=1, relief = 'groove',command=self.selectPath)
        self.clickButton.place(x=700, y=145)

        # 进度和信息展示栏
        # 展示栏title位
        self.titleBar = tk.Frame(self.master, width=777, height=30, bg='#EFEEEE')
        self.titleBar.pack()
        # title
        self.title=tk.Label(self.titleBar, text='运行进度：', bg='#EFEEEE', fg='#575757', font=font)
        self.title.place(x=20,y=0)
        # 清屏按钮
        self.clearButton = tk.Button(self.titleBar, text='清屏',font=('微软雅黑', 11, 'bold'), bg = '#B0C4DE', fg = 'white', activebackground = '#87CEFA', activeforeground = 'white', width=5, height=1, relief = 'groove',command= lambda: self.detail_Text.delete(1.0, "end"))
        self.clearButton.place(x=700, y=0)

        # 信息展示栏
        self.detail_frame = tk.Frame(self.master, width=712, height=139, bg='#F6F7F5')
        self.detail_frame.pack() #place(x=19, y=370)
        scroll = tk.Scrollbar(self.detail_frame)
        self.detail_Text = tk.Text(self.detail_frame, font=('微软雅黑', 11), bg='#F6F7F5', fg='#3E3E3E', yscrollcommand=scroll.set,height=12)
        # 鼠标滚动控件
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.detail_Text.pack(side=tk.LEFT, fill=tk.Y)
        scroll.config(command=self.detail_Text.yview)
        self.detail_Text.config(yscrollcommand=scroll.set)

        # 底部横栏
        self.down_frame = tk.Frame(self.master, width=777, height=45, bg='#EFEEEE')
        self.down_frame.place(x=0, y=510)
        self.button_run = tk.Button(self.down_frame,  text='运行', font=font, bg='#007FFB', fg='white',activebackground='#369AFD', activeforeground='white', width=8, relief='groove')
        self.button_run.place(x=670, y=4)

# 因为转换函数需要在界面输出转换进度和提示信息,所以改成在转换项目中调用gui,以及gui中的print_text
if __name__=="__main__":
    # 创建一个MainUI对象
    app = MainUI()
    # 设置窗口标题
    app.master.title('科学计算器')
    # 设置窗体大小
    app.master.geometry('777x555+500+250')
    app.master.resizable(False, False)
    # 主循环开始
    app.mainloop()
