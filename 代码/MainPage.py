from tkinter import *
from view import *  # 菜单栏对应的各个子页面
import tkinter as tk
from PIL import ImageTk, Image


class MainPage(object):
    def __init__(self, master=None):
        self.root = master  # 定义内部变量root
        self.createPage()   # 创建菜单栏

    def createPage(self):
        self.inputPage = InputFrame(self.root)  # 创建不同Frame
        self.deletePage = DeleteFrame(self.root)
        self.modifyPage = ModifyFrame(self.root)
        self.queryPage = QueryFrame(self.root)
        self.mergePage = MergeFrame(self.root)
        self.visPage = VisFrame(self.root)
        self.stuPage = StuFrame(self.root)
        self.helpPage = HelpFrame(self.root)
        self.inputPage.pack()                  # 默认显示数据录入界面
        menubar = Menu(self.root)
        menubar.add_command(label='增', command=self.inputData)  # 设置不同菜单栏的名字属性
        menubar.add_command(label='删', command=self.deleteData)
        menubar.add_command(label='改', command=self.modifyData)
        menubar.add_command(label='查', command=self.queryData)
        menubar.add_command(label='合', command=self.mergeData)
        menubar.add_command(label='可视化', command=self.visData)
        menubar.add_command(label='个人信息表', command=self.stuData)
        menubar.add_command(label='帮助', command=self.helpData)
        self.root['menu'] = menubar                             # 设置菜单栏

    def inputData(self):                                        # 设置只获取inputData的相关数据
        self.inputPage.pack()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.modifyPage.pack_forget()
        self.mergePage.pack_forget()
        self.visPage.pack_forget()
        self.stuPage.pack_forget()
        self.helpPage.pack_forget()

    def deleteData(self):                                       # 设置只获取deleteData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack()
        self.modifyPage.pack_forget()
        self.mergePage.pack_forget()
        self.visPage.pack_forget()
        self.stuPage.pack_forget()
        self.helpPage.pack_forget()

    def modifyData(self):                                       # 设置只获取modifyData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.modifyPage.pack()
        self.mergePage.pack_forget()
        self.visPage.pack_forget()
        self.stuPage.pack_forget()
        self.helpPage.pack_forget()

    def queryData(self):                                        # 设置只获取queryData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack()
        self.deletePage.pack_forget()
        self.modifyPage.pack_forget()
        self.mergePage.pack_forget()
        self.visPage.pack_forget()
        self.stuPage.pack_forget()
        self.helpPage.pack_forget()

    def mergeData(self):                                          # 设置只获取mergeData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.modifyPage.pack_forget()
        self.mergePage.pack()
        self.visPage.pack_forget()
        self.stuPage.pack_forget()
        self.helpPage.pack_forget()

    def visData(self):                                            # 设置只获取visData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.modifyPage.pack_forget()
        self.mergePage.pack_forget()
        self.visPage.pack()
        self.stuPage.pack_forget()
        self.helpPage.pack_forget()

    def stuData(self):                                           # 设置只获取stuData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.modifyPage.pack_forget()
        self.mergePage.pack_forget()
        self.visPage.pack_forget()
        self.stuPage.pack()
        self.helpPage.pack_forget()

    def helpData(self):                                          # 设置只获取helpData的相关数据
        self.inputPage.pack_forget()
        self.queryPage.pack_forget()
        self.deletePage.pack_forget()
        self.modifyPage.pack_forget()
        self.mergePage.pack_forget()
        self.visPage.pack_forget()
        self.stuPage.pack_forget()
        self.helpPage.pack()