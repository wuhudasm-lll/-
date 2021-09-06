from tkinter import *
from MainPage import *
from tkinter.messagebox import *
import tkinter
import tkinter as tk
from PIL import ImageTk, Image
from tkinter import *
from tkinter import ttk
import os
import time
import tkinter.font as tkFont
import threading
import tkinter as tk
from PIL import ImageTk, Image


def gettime():
    timestr = time.strftime("%H:%M:%S")     # 获取当前时间并转换为字符串
    lb.configure(text=timestr)              # 重新设置标签文本
    root.after(1000, gettime)               # 获取时间


def showWelcome():
    sw = root1.winfo_screenwidth()        # 获取屏幕宽度
    sh = root1.winfo_screenheight()       # 获取屏幕高度r
    root1.overrideredirect(True)          # 去除窗口边框
    root1.attributes("-alpha", 1)         # 窗口透明度（1为不透明，0为全透明）
    x = (sw - 800) / 2
    y = (sh - 450) / 2
    root1.geometry("800x450+%d+%d" % (x, y))      # 将窗口置于屏幕中央
    if os.path.exists(r'C:\Users\余\Desktop\dist\9.gif'):                   # 搜索图片文件（只能是gif格式）
        bm = PhotoImage(file=r'C:\Users\余\Desktop\dist\9.gif')
        lb_welcomelogo = Label(root1, image=bm)      # 将图片放置于窗口
        lb_welcomelogo.bm = bm
        lb_welcomelogo.place(x=-2, y=-2, )           # 设置图片位置


def closeWelcome():
    for i in range(2):
        root.attributes("-alpha", 0)   # 窗口透明度
        time.sleep(1)                  # 屏幕停留时间
    root.attributes("-alpha", 1)       # 窗口透明度
    root1.destroy()                    # 屏幕结束


root = Tk()  # 创建应用程序主窗口
root.title('学生成绩管理系统')  # 设置窗口名字
lb = tkinter.Label(root, text='', fg='dodgerblue', font=("微软雅黑", 12))        # 设置时间的字体和颜色
lb.pack(side=RIGHT, anchor=N)                            # 设置时间的位置
gettime()
root.attributes("-alpha", 0)          # 透明状态下加载主程序
MainPage(root)                        # 进行MainPage.py的程序
msw = root.winfo_screenwidth()        # 获取屏幕宽度
msh = root.winfo_screenheight()       # 获取屏幕高度
m_x = (msw-900)/2
m_y = (msh-500)/2
root.geometry("900x500+%d+%d" % (m_x, m_y))        # 设置主程序窗口置于屏幕中央
global root1                                     # 声明root1为全局变量
root1 = tkinter.Toplevel()                       # 设置欢迎界面的窗口
root1.attributes("-alpha", 0)                    # 设置透明状态全透明
tMain = threading.Thread(target=showWelcome)     # 开始展示
tMain.start()
t1 = threading.Thread(target=closeWelcome)       # 结束展示
t1.start()
root.mainloop()                                  # 窗口循环
