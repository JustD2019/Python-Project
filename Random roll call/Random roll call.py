from tkinter import messagebox,ttk
from tkinter import *
import pandas as pd
import random
import tkinter as tk
import tkinter.filedialog


root = Tk()
root.title("Excel和csv文件随机点名—JustD2019")
root.geometry("380x180+500+300")    #创建窗体并设置初始弹出位置
root.resizable(width=False,height=False)

#存储文件路径类
class save_path:
    f_path = ''
#点名函数
def Random(*event):
    try:
        string1 = int(str1.get())

        if Format.get() == 'xlsx':
            data = pd.read_excel(save_path.f_path, encoding='gbk', header=eval(hed.get()))
        elif Format.get() == 'csv':
            data = pd.read_csv(save_path.f_path, encoding='gbk', header=eval(hed.get()))

        #统计点名列存在的数据个数
        data2 = data.iloc[:, col.get()-1]
        randomData = []
        count_data = 0
        for i in data2:
            randomData.append(i)
            count_data += 1
        #点名提示
        if string1 > count_data:
            messagebox.showerror('错误', '您输入的点名人数大于存在人数! 请重新输入!')
            str1.set('')
        elif string1 <= 0:
            messagebox.showerror('错误', '您输入的点名人数小于等于0! 请重新输入!')
            str1.set('')
        else:
            messagebox.showinfo("点名结果", random.sample(randomData, string1))
    except Exception:
        messagebox.showerror('错误','未选取文件或文件格式或点名人数错误!')
root.bind('<Return>',Random)   #绑定回车键


#清零数据个数并打开文件并把文件路径传给类
def OpenFile():
    save_path.data_count = 0
    file_path = tkinter.filedialog.askopenfilename()  # 返回文件名
    save_path.f_path = file_path


#创建文件、选择文件并且填写数据
def CreateFile():
    save_path.data_count = 0
    messagebox.showinfo('提示','仅提供创建csv文件的存储功能!')
    tkinter.filedialog.asksaveasfile()#创建文件
    file_path = tkinter.filedialog.askopenfilename()#返回文件名
    save_path.f_path = file_path
    # ***********************************************************************************************#
    root2 = tk.Toplevel(root)
    root2.title("存储数据")
    root2.geometry("380x180+500+300")
    root2.config(bg="white")
    root2.resizable(width=False, height=False)

    def save(*event):
        data = save_data.get()
        data = pd.DataFrame({data})   #存储DataFrame数据，存储于第一列
        data.to_csv(file_path, encoding='gbk', header=None, mode='a', index=False)
        save_data.set('')
    root2.bind('<Return>',save)
    def read():
        r_data1 = pd.read_csv(file_path, encoding='gbk', header=None)
        r_data2 = r_data1.iloc[:, 0]
        read_data = []
        for i in r_data2:
            read_data.append(i)
        root2.withdraw()  #隐藏窗体
        messagebox.showinfo('已存取数据',read_data)
        root2.wm_deiconify() #显示窗体

    def save_end():
        root2.withdraw()  # 存储完成隐藏窗体

    # ***********************************************************************************************#
    save_l1 = Label(root2, text="数据:", font=("宋体", 15), bg="white")
    save_l1.place(x=60, y=40)
    # ***********************************************************************************************#
    save_btn1 = Button(root2, text="存储", font=("宋体", 15), bg="white", command=save)
    save_btn1.place(x=80, y=100)

    save_btn2 = Button(root2, text="查看", font=("宋体", 15), bg="white", command=read)
    save_btn2.place(x=150, y=100)

    save_btn3 = Button(root2, text="存储结束", font=("宋体", 15), bg="white", command=save_end)
    save_btn3.place(x=220, y=100)
    # ***********************************************************************************************#
    save_data = StringVar()
    save_entry1 = Entry(root2, width="20", textvariable=save_data)
    save_entry1.place(x=120, y=40)
    #***********************************************************************************************#


#***********************************************************************************************#
#用单选按钮来进行选择
Format = StringVar()
radio1 = tk.Radiobutton(root,text='xlsw格式',variable=Format,value='xlsx',font=("微软雅黑",12))
radio1.place(x='50',y='80')
radio1.select() #默认选中的单选框

radio2 = tk.Radiobutton(root,text='csv格式',variable=Format,value='csv',font=("微软雅黑",12))
radio2.place(x='180',y='80')

hed = StringVar()
radio3 = tk.Radiobutton(root,text='是',variable=hed,value='0',font=("微软雅黑",12))
radio3.place(x='50',y='20')
radio3.select() #默认选中的单选框

radio4 = tk.Radiobutton(root,text='否',variable=hed,value='None',font=("微软雅黑",12))
radio4.place(x='180',y='20')

#***********************************************************************************************#
l1 = Label(root,text="点名人数: ",font=("微软雅黑",13))
l1.place(x='40',y='50')

l2 = Label(root,text="点名列",font=("微软雅黑",8))
l2.place(x='280',y='33')

l3 = Label(root,text="点名列是否有列名 ",font=("微软雅黑",12))
l3.place(x='40',y='3')

#***********************************************************************************************#
#创建下拉选择框来选择点名的列数
col = IntVar()
cmb = ttk.Combobox(root,textvariable=col,width=3)
cmb.place(x='280',y='50')
s = []
for i in range(1,200):
    s.append(i)
cmb['value']=(s)  #下拉选择框里的值
cmb.current(0)    #默认选择第一个数据

#***********************************************************************************************#

str1 = StringVar()
entry1 = Entry(root,textvariable=str1,width=15,bd=5)
entry1.place(x='140',y='50')

#***********************************************************************************************#

btn1 = Button(root,text="点名",font=("微软雅黑",13),width="7",command=Random)
btn1.place(x='230',y='120')

btn2 = Button(root,text="创建文件",font=("微软雅黑",13),width="7",command=CreateFile)
btn2.place(x='130',y='120')

btn3 = Button(root,text="打开文件",font=("微软雅黑",13),width="7",command=OpenFile)
btn3.place(x='30',y='120')

#***********************************************************************************************#
root.mainloop()