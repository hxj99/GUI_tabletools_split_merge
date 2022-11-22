#!/usr/bin/python
# encoding: utf-8

import tkinter.ttk as ttk
import tkinter as tk
from tkinter import filedialog
import pandas as pd ,re,os,datetime,time
import tkinter.messagebox as msgbox
import tkinter.simpledialog as inputbox

# https://blog.csdn.net/weixin_45558166/article/details/122079072
# 可编辑的表格
# v1.2.5主要调整修改时弹出框为simpledialog

class MyTable(ttk.Frame):
    
    def __init__(self, master):
        self.master = master
        super().__init__(master)
        
        self.master.geometry('800x400+400+200')
        self.master.title("Excel表格工具v1.0.0")
        self.pack()              #frame加入root
        # 标题
        self.toptitle = tk.Label(self, text='Excel表格合并与拆分',font=('仿宋',14,'bold'))
        self.toptitle.pack(pady=5)
        
        self.table_frame = ttk.Frame(self.master)   # 标题区域
        self.btn_frame = tk.Frame(self.master)      # 列表区域
        self.bottom_frame = tk.Frame(self.master)   # 底部区域
        self.table_frame.pack()
        self.btn_frame.pack()
        self.bottom_frame.pack()
        
        tableColumns = ['*excel文件名称', '*工作表', '*开始行号','结束行号','取出的列A,B,C...或A:C','关键列ABC...']
        tableColumnsWifth = [350,80,60,60,150,100]
        # tableValues = [
                    # ["E:\MyCode\自制软件工具\测试数据\自定义测试数据.xlsx", "1","2","","",""],
                    # ["E:\MyCode\自制软件工具\测试数据\自定义测试数据.xlsx", "2","4","","",""],
        # ]
        # 设置滚动条
        self.xscroll = tk.Scrollbar(self.table_frame, orient='horizontal', command=self.on_X_ScrollCommand)
        self.yscroll = tk.Scrollbar(self.table_frame, orient='vertical', command=self.on_Y_ScrollCommand)
        self.xscroll.pack(side='bottom', fill='x')
        self.yscroll.pack(side='right', fill='y')
        # 修改提示
        tk.Label(self.table_frame, text='注：双击内容进行修改。').pack(side="bottom",fill='x')
        
        self.table = ttk.Treeview(
            master=self.table_frame,        # 父容器
            columns=tableColumns,           # 列标识符列表（标题行）
            height=5,                       # 表格显示的行数
            show='headings',                # 隐藏首列
            style='Treeview',               # 样式
            displaycolumns="#all",          # 列样式
            selectmode='browse',            # 单选
            xscrollcommand=self.xscroll.set,     # x轴滚动条
            yscrollcommand=self.yscroll.set      # y轴滚动条
        )

        self.table.pack()  #TreeView加入frame

        # 设置标题栏和列
        for i in range(len(tableColumns)):
            self.table.heading(column=tableColumns[i], text=tableColumns[i], anchor='center') # 设置表头
            self.table.column(tableColumns[i], width=tableColumnsWifth[i], anchor='center', minwidth=60,stretch=True)    # 设置列

        style = ttk.Style(self.master)
        style.configure('Treeview', rowheight=30)
        
        # 添加数据内容
        # for i, data in enumerate(tableValues):
            # self.table.insert('', 'end', value=data)  #, tags=str(i)
        
        # 表格绑定事件
        self.table.bind("<Double-1>",self.on_Double_1)          # 左键双击
        
        
        # 底部添加按钮,frame2
        self.addbtn = tk.Button(self.btn_frame, text='添加',width=10, command=self.addbtn,bg='#5882FA',font=('仿宋',14,'bold'))
        self.deletebtn = tk.Button(self.btn_frame, text='删除',width=10, command=self.deletebtn,bg='red',font=('仿宋',14,'bold'))
        self.mergebtn = tk.Button(self.btn_frame, text='合并',width=10, command=self.mergebtn,bg='#04B404',font=('仿宋',14,'bold'))
        self.modifybtn = tk.Button(self.btn_frame, text='拆分',width=10,command=self.splitbtn,bg='#FAAC58',font=('仿宋',14,'bold'))
        self.addbtn.pack(side="left",padx=10,pady=40)
        self.deletebtn.pack(side="left",padx=10,pady=40)
        self.mergebtn.pack(side="left",padx=10,pady=40)
        self.modifybtn.pack(side="left",padx=10,pady=40)
        

        #底部自选输出路径
        self.outdir_var = tk.StringVar()
        self.outdir_label = tk.Label(self.bottom_frame, text='保存位置：')
        self.outdir_Entry = tk.Entry(self.bottom_frame, textvariable=self.outdir_var,width=90,state='disabled')    # 输入内容
        self.outdir_selectbtn = tk.Button(self.bottom_frame, text='选择文件夹',width=10, command=self.outdirbtn)
        self.outdir_label.pack(side="left")
        self.outdir_Entry.pack(side="left")
        self.outdir_selectbtn.pack(side="left")

    # 自定义提示信息
    def msgbox(self,type,msg):
        if type=="info":
            msgbox.showinfo(title = '提示', message=msg)
        elif type=="warning":
            msgbox.showwarning(title = '提示', message=msg)
        elif type=="error":
            msgbox.showerror(title = '提示', message=msg)
        else:
            print("类型参数错误")
           
    # 自定义询问框
    def askbox(self,type,msg):
        if type=="okcancel":      # 确定，取消
            yn = msgbox.askokcancel(title = '提示', message=msg)   # 返回值为True或者False
        elif type=="question":      # 是，否
            yn = msgbox.askquestion(title = '提示', message=msg)   # 返回值为：yes/no
        elif type=="retrycancel":   # 重试，取消
            yn =msgbox.askretrycancel(title = '提示', message=msg) # 返回值为：True或者False
        elif type=="yesno":           # 是，否
            yn = msgbox.askyesno(title = '提示', message=msg)      # 返回值为True或者False
        else:
            print("类型参数错误")
            yn = "TypeError"
        return yn
    
    # 自定义用户输入框
    def userinputbox(self,type,msg,initialvalue):
        # title = '录入信息',prompt='请输入姓名：',initialvalue = '可以设置初始值'
        if type=="string":          # 获取字符串
            userinput = inputbox.askstring(title = '录入信息',prompt=msg,initialvalue = initialvalue)
        elif type=="integer":       # 获取整数
            userinput = inputbox.askinteger(title = '录入信息',prompt=msg,initialvalue = initialvalue)
        elif type=="float":         # 获取浮点数
            userinput = inputbox.askfloat(title = '录入信息',prompt=msg,initialvalue = initialvalue)
        else:
            print("类型参数错误")
            userinput = "TypeError"
            
        return userinput
        
    # 鼠标左键双击触发
    def on_Double_1(self, event):
        # print("左键双击",event.widget)
        if str(event.widget) == ".!frame.!treeview":  # 双击触发的是否为表格(根据event.widget判断)

            table = event.widget

            self.row = table.identify_row(event.y)         # 点击的行，当点击标题行时，内空为空
            self.column = table.identify_column(event.x)   # 点击的列，形式#1，#2，#3
            
            # print("双击行和列：",self.row,self.column)
            if not self.row:
                # print("表格标题,不能修改！")
                return
            if self.column =="#1":
                # print("第一列不能修改")
                return
            col = int(str(table.identify_column(event.x)).replace('#', ''))  # 列号取出数字部分
            text = table.item(self.row, 'value')[col - 1]    # 单元格内容

            tagdict = {'#2':'请输入：工作表（数字）','#3':'请输入：开始行号（数字）','#4':'请输入：结束行号（数字）',
                       '#5':'请输入：取出的列（字母）','#6':'请输入：关键列（字母）',}
            texttypedict = {'#2':'N2','#3':'N5','#4':'N5','#5':'A99','#6':'A1'}
            taginfo = tagdict[self.column]
            userinput = self.userinputbox("string",taginfo,text)
            # print("用户输入：",userinput)
            if userinput==None:
                return
            if len(userinput)>0:
                # 1）输入内容进行正则匹配
                t = texttypedict[self.column]
                res_reg = self.textregexps(userinput,t)
                if res_reg==False:
                    # print("输入内容的格式错误")
                    self.msgbox("error","格式错误")
                    return
                # 2）正式修改输入内容
                self.table.set(self.row, self.column, userinput) # 表格数据设置为输入条内容
                # print("修改完成！")
            else:
                if self.column in ['#4','#5','#6']:
                    yn = self.askbox("okcancel","你确定要输入空值？")
                    if yn:
                        self.table.set(self.row, self.column, userinput)
                        # print("修改完成！")
                    else:
                        return
                else:
                    self.msgbox("error","不能输入空值！")
                    return
                    
    # 正则表达式校验，返回TRUE/FALSE,t为校验类别
    def textregexps(self,s,t):
        # t的类型有: A工作表，B行号，C取出的列
        if t=="N2" and re.match(r"^[1-9][0-9]{0,1}$",s):         # 工作表索引 0-99
            return True
        elif t=="N5" and re.match(r"^[1-9][0-9]{0,4}$",s):       # 行号 0-99999 十万条
            return True
        elif t=="A99" and re.match(r"^([a-zA-Z]{1,2}[,:]?)+$",s): # 取出的列（a,b,c...; A:C）
            return True
        elif t=="A1" and re.match(r"^[a-zA-Z]+$",s):    # 关键的列
            return True
        else:
            return False
    
    # 进一步校验列表内容并转换成二维数组 ；文本转数字,items是二维数组
    def checkitems(self,items):        
        tableArr = []
        for i,data in enumerate(items):
            # print("条目信息：\n",data)
            # (1)文件名
            name = data[0]                  
            # (2)工作表，仅输入数字
            shtname = data[1]
            sheetname = int(data[1])-1      # 工作表名称（默认0）

            # (3)开始行号
            startrow = int(data[2])-1
            # (4)结束行号
            endrow = data[3]
            if endrow:
                endrow = int(endrow) - 1       # 结束行号
                if startrow>endrow:
                    msg = "'结束行号'不能大于'开始行号'！"
                    return False,msg
            else:
                endrow = None
            # (5)取出的列
            columns = data[4]       # 合并，取出的列
            # (6)关键的列
            keycolumn = data[5]     # 拆分，关键的列

            # 重新生成标准的列表信息(行号文本转换为数字)。
            arr = []
            arr = [name,sheetname,startrow,endrow,columns,keycolumn]
            tableArr.append(arr)

        return True,tableArr
 

    # X轴滚动条拖动触发
    def on_X_ScrollCommand(self, *xx):
        self.table.xview(*xx)          # 表格横向滚动
        
    # Y轴滚动条拖动触发
    def on_Y_ScrollCommand(self, *xx):
        self.table.yview(*xx)          # 表格纵向滚动
        

    # 添加一行
    def addbtn(self):
        # print("添加一行Btn")
        # 打开文件选择框,  # askopenfilenames函数选择多个文件,返回的是元组，绝对路径
        selected_files = filedialog.askopenfilenames(title="请选择Excel文件",filetypes=[("Excel","*.xlsx;*.xls")])
        # print(selected_files)

        #插入一行
        for i, data in enumerate(selected_files):
            self.table.insert('', 'end', value=(data,1,2,"","",""))
        # self.msgbox("info","添加完成！")

        
    # 删除一行
    def deletebtn(self):
        # print("删除一行Btn")
        # 取得选中的行,返回的是元组(I001,I002)
        index = self.table.selection()
        # print(index)
        if index:
            yn = self.askbox("okcancel","你确定要删除？")
            # print(yn)
            if yn:
                self.table.delete(index)
                # print("删除成功")
                # self.msgbox('info','删除完成')
        else:
            self.msgbox("warning","请选择要删除的文件或工作表！")
        
    # 选择输出的文件夹Btn
    def outdirbtn(self):
        selected_dir = filedialog.askdirectory(title="选择文件夹")  # ,initialdir='D:/'
        # print(selected_dir)
        if selected_dir:
            self.outdir_var.set(selected_dir)
    
    
    # 合并表格Btn(处理表格条目)   
    def mergebtn(self):
        # print("合并表格Btn")
        # 输出路径是否为空
        outdir = self.outdir_var.get()
        if len(outdir)<3:
            # print("保存位置为空")
            self.msgbox("warning",'保存位置不能为空，请选择一个文件夹！')
            return

        # 1.获取列表（条目）信息，二维数组[[],[]]
        items = [list(self.table.item(x,"values")) for x in self.table.get_children()]
        # print("列表内容：\n",items)
        # 判断是否工作表数量，当一个表格时不能进行合并操作。
        if len(items)<2:
            # print("至少2个工作表才能进行合并操作！")
            self.msgbox("warning",'至少2个工作表,才能合并！')
            return
        
        # 2.列表（条目）信息校验
        checkresult,tableArr = self.checkitems(items)
        
        if not checkresult:  # 校验未通过
            # print(tableArr)
            self.msgbox("warning",tableArr)
            return
        # 读取多个工作表
        arr = self.readtabledata(tableArr)
        # 合并处理
        df = self.moretableconcat(arr)
        # 保存表格
        # file = items[0][0]
        
        self.savetable(outdir,df,'join')
        self.msgbox("info",'合并完成！')
            
    # pandas打开工作表读取数据,返回二维数组
    def readtabledata(self,items):
        tbArr = []
        for item in items:
            # print(item)
            file = item[0]
            sheetname = item[1]
            skiprows = item[2]-1    # 跳过前面多少行(不包含标题行，所减1)，开始行号item[2]
            endrow = item[3]
            columns = item[4]       # 指定列，数字或者字母
            keycolumn = item[5]     # 拆分关键的列
            
            # 开始行号与标题行号联动关系
            if skiprows<0 :
                tagrow=None
            else:
                tagrow=0
            # 结束行号
            if endrow:
                nrows = endrow - skiprows       # 取多少行(从标题行开始到结束行)
            else:
                nrows = None
            # 取得指定的列
            if columns:
                usecols = columns
            else:
                usecols = None
            
            # print("read_excel参数：\n",file,sheetname,tagrow,skiprows,usecols)
            
            with open(file, 'rb') as f:
                tb = pd.read_excel(f,sheet_name=sheetname,header=tagrow,skiprows=skiprows, nrows=nrows, usecols=usecols)
                tbArr.append(tb)
            
            # print("读取表格数据成功\n",tb.info(),"\n")

            # print("行数：",len(tb))
            
        return tbArr
    
    
    # 多个表格进行合并处理
    def moretableconcat(self,arr):
        df = pd.concat(arr, axis=0, join='outer', ignore_index=True)
        # print("多表合并后信息：\n")
        # print(df.info())
        # print("-"*20,"\n\t行数和列数",df.shape)
        # print("-"*20,"前3行\n",df.head(3))
        # print("-"*20,"尾3行\n",df.tail(3))
        
        # print("合并完成!")
        
        return df

    # 保存表格
    def savetable(self,dir,df,type,key=None):
        # type=join,split
        # 文件路径和名称
        # dir = os.path.dirname(file)
        # name = os.path.basename(file)
        
        curr_time = datetime.datetime.now()
        time_str = datetime.datetime.strftime(curr_time,'%Y%m%d-%H%M%S')
        if type=='join':
            newfilename = "合并表格_" + time_str + ".xlsx"
        elif type=='split':
            newfilename = key +"_" +time_str + ".xlsx"
        else:
            print("保存函数参数错误")
            return
        # 合成全新保存路径和名称
        savefile = os.path.join(dir,newfilename)    
        df.to_excel(savefile,index=False,encoding='uft-8')
        print(savefile,"\n保存完成！")
        

    
    # 按关键列拆分
    def keyssplittable(self,df,column):
        columnnumber = ord(column)
        # print("字母转数字：",columnnumber)
        if columnnumber<91:  # A-Z:65-90,a-z:97-122
            n = columnnumber-65
            keycolumn = df.iloc[:,n]  # 列索引从0开始
        else:
            n = columnnumber-97
            keycolumn = df.iloc[:,n]
        listkey = keycolumn.tolist()   
        keysArr = list(set(listkey))   # 去重
        # print("关键字去重完成!\n",keysArr)   

        dfArr = []
        for key in keysArr:
            dfkey = df[df.iloc[:,n]==key]    
            dfArr.append(dfkey)
            # print(key,dfkey.shape)
    
        return dfArr,keysArr
    
    
    
    # 拆分按钮
    def splitbtn(self):
        # print("拆分Btn")
        # 输出路径是否为空
        outdir = self.outdir_var.get()
        if len(outdir)<3:
            # print("保存位置为空")
            self.msgbox("warning",'保存位置不能为空，请选择一个文件夹！')
            return
        
        
        # 1.取得选中的行，返回I001
        index = self.table.selection()
        if len(index)<1:
            # print("请选择一个文件")
            self.msgbox("warning","请在列表中选择一行！")
            return
           
            
        item = self.table.item(index,"values")
        # print(item)
        itemarr = [list(item)]   # 元组转列表,并生成二维数组 
        if len(itemarr[0][5])<1:
            # print("请输入关键的列ABC..")
            self.msgbox("warning","请输入关键列ABC..")
            return
        # print(arr,arr[-1])
        
        # 2.校验条目数据并转换格式（文本为数字）
        checkresult,msgorarr = self.checkitems(itemarr)
        if not checkresult:
            # print(msgorarr)
            return
        
        # 3.读取文件表格，返回数据
        # print(msgorarr)
        df = self.readtabledata(msgorarr)[0]  # 函数返回的是二维数组，拆分只有一个表。
        # print(df.head(3))
        # print(df.tail(3))
        # print("读取文件完成！")
        
        # 4.按关键列分别取出,返回列表
        column = itemarr[0][5]
        dfArr,keysArr = self.keyssplittable(df,column)
        
        # 5.保存文件
        # print("保存文件完成！")
        # file = itemarr[0][0]
        for i,df in enumerate(dfArr):
            key = keysArr[i]
            self.savetable(outdir,df,'split',key)
        
        # print("拆分完成！")
        self.msgbox("info","拆分完成！")


        
if __name__ == '__main__':
    root = tk.Tk()
    MyTable(root)
    root.mainloop()