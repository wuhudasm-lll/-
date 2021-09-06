from tkinter import messagebox
from tkinter import *
from tkinter import ttk
import xlwt
import xlrd
from xlutils.copy import copy
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from MainPage import *
import webbrowser


def set_style(name, height, borders, colour, colour_index):   # 设置写入表格的字体和格式
    style = xlwt.XFStyle()
    style.font = xlwt.Font()
    style.font.name = name  # 设置字体
    style.font.height = height  # 设置字体大小
    style.font.colour_index = colour_index  # 设置字体颜色
    style.borders = xlwt.Borders()
    # 设置边框
    style.borders.left = borders
    style.borders.top = borders
    style.borders.right = borders
    style.borders.bottom = borders
    # 设置背景
    style.pattern = xlwt.Pattern()
    style.pattern.pattern = True  # 允许设置背景
    # style.pattern.pattern_back_colour = 0x04    # 后背景颜色
    style.pattern.pattern_fore_colour = colour  # 前背景颜色
    # 设置表格对齐属性
    style.alignment = xlwt.Alignment()
    style.alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平对齐
    style.alignment.wrap = xlwt.Alignment.NOT_WRAP_AT_RIGHT  # 自动换行
    style.alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直对齐
    return style


class InputFrame(Frame):  # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master  # 定义内部变量root
        self.E1 = Entry(self)  # 设置得到输入的内容
        self.E2 = Entry(self)
        self.E3 = Entry(self)
        self.E4 = Entry(self)
        self.E5 = Entry(self)
        self.E6 = Entry(self)
        self.E7 = Entry(self)
        self.E8 = Entry(self)
        self.E9 = Entry(self)
        self.E10 = Entry(self)
        self.E11 = Entry(self)
        self.E12 = Entry(self)
        self.E13 = Entry(self)
        self.E14 = Entry(self)
        self.createPage()

    def Isspace(self, text):
        temp = 0                           # 检索text中的内容是否为空，如果是赋值为1
        for i in text:
            if not i.isspace():
                temp = 1
                break
        if temp == 1:
            return 0
        else:
            return 1

    def write(self, num, name, major, calss, score1,score2,score3,score4,score5,score6,score7,score8,score9,score10):
        workbook = xlrd.open_workbook("学生信息成绩表.xls")                                         # 打开表格
        data = xlrd.open_workbook('学生信息成绩表.xls', formatting_info=True)
        sheet = workbook.sheet_by_index(0)                                                         # 获取表格所需要的页
        my_dict = dict()                                                                           # 定义一个字典
        for i in range(2, sheet.nrows):                                                            # 通过循环遍历从第三行行到结束
            my_dict_key = sheet.cell_value(i, 1)                                                   # 定义每行第二列的数据为字典的key值
            my_dict_value = sheet.cell_value(i, 2)                                                 # 定义每行第三列的数据为字典的value
            my_dict[my_dict_key] = my_dict_value                                                   # 定义key，和value
        my_dict = list(map(int, my_dict))                                                          # 将字典里的数据转化为整形
        excel = copy(wb=data)                                                                      # 完成xlrd对象向xlwt对象转换
        excel_table = excel.get_sheet(0)
        values = (num, name, major, calss, score1, score2, score3, score4, score5, score6, score7, score8, score9, score10)
        i = 0
        leap = 0
        for temp in my_dict:                                     # 循环遍历字典
            if temp == int(num):                                 # 如果循环到的数据等于输入的学号
                leap = 1
                break
            else:
                i = i + 1
        while leap == 1:                                         # 如果leap = 1那么就输出不存在该学生信息
            messagebox.showinfo(title='结果', message="已存在该学生科目信息！")
            break
        else:
            ncols = 1
            for value in values:                                # 循环遍历输入的数据
               excel_table.write(sheet.nrows, ncols, value, set_style('Courier New', 20 * 9, 0x01, 42, 48))          # 循环遍历写入表格，并设置格式
               ncols = ncols + 1                                # 行数加一
            messagebox.showinfo(title='提示', message="写入成功")
            excel_table.write(sheet.nrows, 0, sheet.nrows - 1, set_style('Courier New', 20 * 9, 0x01, 42, 48))       # 写入序号
        excel.save('学生信息成绩表.xls')                         # 保存表格
        return

    def click(self):
        num = self.E1.get()                                          # 得到输入的数据
        name = self.E2.get()
        major = self.E3.get()
        calss = self.E4.get()
        score1 = self.E5.get()
        score2 = self.E6.get()
        score3 = self.E7.get()
        score4 = self.E8.get()
        score5 = self.E9.get()
        score6 = self.E10.get()
        score7 = self.E11.get()
        score8 = self.E12.get()
        score9 = self.E13.get()
        score10 = self.E14.get()
        if self.Isspace(num) or self.Isspace(name) or self.Isspace(major) or self.Isspace(calss):    # 判断输入的学号，名字，班级专业是否为空
            messagebox.showinfo(title='提示', message="输入项为空")
        else:                                                             # 如果不是，则写入输入的数据
            self.write(num, name, major, calss, score1, score2, score3, score4, score5, score6, score7, score8, score9, score10)

    def createPage(self):           # 设置文本和文本框，以及按钮组件
        Label(self).grid(row=0, stick=W, pady=10)
        Label(self, text='学号：', font=("华文彩云")).grid(column=0, row=1, stick=W, pady=10)
        self.E1.grid(row=1, column=1, stick=E)
        Label(self, text='姓名：', font=("华文彩云")).grid(column=3, row=1, stick=W, pady=10)
        self.E2.grid(row=1, column=4, stick=E)
        Label(self, text='专业：', font=("华文彩云")).grid(row=3, stick=W, pady=10)
        self.E3.grid(row=3, column=1, stick=E)
        Label(self, text='班级：', font=("华文彩云")).grid(column=3, row=3, stick=W, pady=10)
        self.E4.grid(row=3, column=4, stick=E)
        Label(self, text='高级程序语言：', font=("华文彩云")).grid(row=6, stick=W, pady=10)
        self.E5.grid(row=6, column=1, stick=E)
        Label(self, text='python编程：', font=("华文彩云")).grid(column=3, row=6, stick=W, pady=10)
        self.E6.grid(row=6, column=4, stick=E)
        Label(self, text='数据库原理：', font=("华文彩云")).grid(row=7, stick=W, pady=10)
        self.E7.grid(row=7, column=1, stick=E)
        Label(self, text='数据结构与算法：', font=("华文彩云")).grid(column=3, row=7, stick=W, pady=10)
        self.E8.grid(row=7, column=4, stick=E)
        Label(self, text='数学分析：', font=("华文彩云")).grid(row=9, stick=W, pady=10)
        self.E9.grid(row=9, column=1, stick=E)
        Label(self, text='高等数学：', font=("华文彩云"), fg="black").grid(column=3, row=9, stick=W, pady=10)
        self.E10.grid(row=9, column=4, stick=E)
        Label(self, text='网络爬虫：', font=("华文彩云")).grid(row=11, stick=W, pady=10)
        self.E11.grid(row=11, column=1, stick=E)
        Label(self, text='数据可视化：', font=("华文彩云")).grid(column=3, row=11, stick=W, pady=10)
        self.E12.grid(row=11, column=4, stick=E)
        Label(self, text='数据挖掘：', font=("华文彩云")).grid(row=13, stick=W, pady=10)
        self.E13.grid(row=13, column=1, stick=E)
        Label(self, text='数据分析：', font=("华文彩云")).grid(column=3, row=13, stick=W, pady=10)
        self.E14.grid(row=13, column=4, stick=E)
        Label(self, text='增加学生信息', font=("华文彩云")).grid(row=0, column=2, stick=N, pady=10)
        Label(self, text='各科成绩', font=("华文彩云")).grid(row=5, column=2, stick=N, pady=10)
        action = ttk.Button(self, text="录入", command=self.click)
        action.grid(row=15, column=4, stick=E, pady=10)


class DeleteFrame(Frame):  # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master  # 定义内部变量root
        self.E1 = Entry(self)              # 定义输入数据
        self.E2 = Entry(self)
        self.createPage()

    def Isspace(self, text):               # 检索text中的内容是否为空，如果是赋值为1
        temp = 0
        for i in text:
            if not i.isspace():
                temp = 1
                break

        if temp == 1:
            return 0
        else:
            return 1

    def click(self):
        df = pd.read_excel('学生信息成绩表.xls')  # 读取Excel表
        df.dropna(how='all', inplace=True)    # 过滤缺失数据，删除全为空值(NaN)的一行并在原表上进行修改
        df.to_excel('学生信息成绩表.xls', sheet_name='Sheet1', index=False, encoding='UTF-8')   # 保存，设置格式
        workbook1 = xlrd.open_workbook('学生信息成绩表.xls')
        # 获取表格信息
        sheet1 = workbook1.sheet_by_name(workbook1.sheet_names()[0])
        rows1 = sheet1.nrows
        cols1 = sheet1.ncols
        workbook = xlwt.Workbook(encoding="UTF-8")
        sheets = workbook.add_sheet("Sheet1")
        # 设置列宽
        sheets.col(0).width = 256 * 5
        sheets.col(1).width = 256 * 18
        sheets.col(2).width = 256 * 18
        for i in range(3, 15):
            sheets.col(i).width = 256 * 14
        # 设置行高
        sheets.row(0).height_mismatch = True  # 允许设置行高
        sheets.row(0).height = 20 * 21
        for i in range(1, rows1):
            sheets.row(i).height_mismatch = True
            sheets.row(i).height = 20 * 15

        def set_styles(name, height, borders, colour, colour_index):  # 设置样式
            styles = xlwt.XFStyle()
            styles.font = xlwt.Font()
            styles.font.name = name  # 设置字体
            styles.font.height = height  # 设置字体大小
            styles.font.colour_index = colour_index  # 设置字体颜色
            styles.borders = xlwt.Borders()
            # 设置边框
            styles.borders.left = borders
            styles.borders.top = borders
            styles.borders.right = borders
            styles.borders.bottom = borders
            # 设置背景
            styles.pattern = xlwt.Pattern()
            styles.pattern.pattern = True  # 允许设置背景
            # style.pattern.pattern_back_colour = 0x04    # 后背景颜色
            styles.pattern.pattern_fore_colour = colour  # 前背景颜色
            # 设置表格对齐属性
            styles.alignment = xlwt.Alignment()
            styles.alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平对齐
            styles.alignment.wrap = xlwt.Alignment.NOT_WRAP_AT_RIGHT  # 自动换行
            styles.alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直对齐
            return styles

        sheets.write_merge(0, 0, 0, 13, '学生信息成绩表', set_styles('黑体', 20 * 16, 0x01, 41, 6))  # 合并单元格写入内容，设置样式
        for i in range(1, rows1):          # 获取行的内容
            for j in range(0, cols1):      # 获取列的内容
                sheets.write(i, j, sheet1.row_values(i)[j], set_styles('Courier New', 20 * 9, 0x01, 42, 48))   # 向指定行列写入信息

        workbook.save("学生信息成绩表.xls")  # 保存
        num = self.E1.get()
        course = self.boxChoice.get()
        if self.Isspace(num) or self.Isspace(course):                      # 如果输入的值为空，就提示输入为空
            messagebox.showinfo(title='提示', message="输入项为空")
        else:
            wb = xlrd.open_workbook("学生信息成绩表.xls", formatting_info=True)        # 打开表格，并保持表格格式不变
            sheet = wb.sheet_by_index(0)                                   # 获取表页
            dic = {}                                                       # 定义一个空字典
            for i in range(2, sheet.nrows):                                # 循环表里第三行到最后一行
                lis = []                                                   # 定义一个列表
                for j in range(1, sheet.ncols):                            # 循环遍历从第一列到最后一列的数据
                    lis.append(sheet.cell(i, j).value)                     # 再将每行每列的数据放入列表中
                dic[sheet.cell(i, 1).value] = lis                          # 将第一列的数据定义为列表的key值
            my_dic = list(map(int, list(dic.keys())))                      # 将字典的值转化为整形
            sum = 0
            temp = 0
            a = 0
            new_wb = copy(wb)                                              # 将原有的Excel，拷贝一个新的副本
            new_sheet = new_wb.get_sheet(0)
            for i in my_dic:                                               # 遍历循环字典里的值
                a = a + 1
                if i == int(num):                                          # 判定指定属性，确定删除行
                    if course == "全部":
                        sum += 1
                        dic.pop(str(num))                                  # 删除输入学号的一行
                        m = 0
                        for i in list(dic.keys()):                         # 循环遍历字典中的key值
                            m += 1
                            n = 1
                            for j in dic[i]:                               # 循环遍历字典中的value
                                new_sheet.write(m+1, n, j, set_style('Courier New', 20 * 9, 0x01, 42, 48))    # 将遍历的值以一种格式输入到表格
                                n += 1
                        for h in range(m + 1, m + 1 + sum):          # 循环遍历
                            m += 1
                            n = 1
                            for k in dic[i]:
                                new_sheet.write(m+1, n, '', set_style('Courier New', 20 * 9, 0x01, 42, 48))        # 将空白的值填入以覆盖之前被删除的值
                                n += 1
                        new_sheet.write(sheet.nrows - 1, 0, '', set_style('Courier New', 20 * 9, 0x01, 42, 48))     # 最后一行的序号值填入空
                        temp = 1
                        break
                    elif course == "高级程序语言":
                        if sheet.cell_value(a+1, 5) == '':                          # 判断是否为空值
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 5, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))    # 如果不是就给予一个空值
                    elif course == "python编程":
                        if sheet.cell_value(a+1, 6) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 6, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "数据库原理":
                        if sheet.cell_value(a+1, 7) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 7, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "数据结构与算法":
                        if sheet.cell_value(a+1, 8) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 8, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "数学分析":
                        if sheet.cell_value(a+1, 9) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 9, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "高等数学":
                        if sheet.cell_value(a+1, 10) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 10, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "网络爬虫":
                        if sheet.cell_value(a+1, 11) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 11, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "数据可视化":
                        if sheet.cell_value(a+1, 12) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 12, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "数据挖掘":
                        if sheet.cell_value(a+1, 13) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 13, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                    elif course == "数据分析":
                        if sheet.cell_value(a+1, 14) == '':
                            messagebox.showinfo(title='提示', message="该科目成绩不存在！")
                            temp = 2
                            break
                        else:
                            temp = 1
                            new_sheet.write(a+1, 14, '—', set_style('Courier New', 20 * 9, 0x01, 42, 48))
                        break
            if temp == 1:                                                    # 判断temp的值
              messagebox.showinfo(title='提示', message="删除成功！")
            elif temp == 0:
              messagebox.showinfo(title='提示', message="不存在该学生信息！")
            new_wb.save("学生信息成绩表.xls")                                 # 保存表格

    def createPage(self):                                                   # 编辑文本，按钮，文本框
        Label(self).grid(row=0, stick=W, pady=10)
        Label(self, text='删除学生信息', font=("华文彩云")).grid(row=0, column=1, stick=N, pady=10)
        Label(self, text='学号: ', font=("华文彩云")).grid(row=1, stick=W, pady=10)
        self.E1.grid(row=1, column=1, stick=E)
        Label(self, text='科目: ', font=("华文彩云")).grid(row=2, stick=W, pady=10)
        sexBoxValue = StringVar()
        self.boxChoice = ttk.Combobox(self, width=17, textvariable=sexBoxValue, state='readonly')
        self.boxChoice['value'] = ("高级程序语言", "python编程", "数据库原理", "数据结构与算法", "数学分析", "高等数学", "网络爬虫", "数据可视化", "数据挖掘", "数据分析", "全部")
        self.boxChoice.current(0)
        self.boxChoice.grid(row=2, column=1, sticky=E)
        action = ttk.Button(self, text="删除", command=self.click)
        action.grid(row=6, column=1, stick=E, pady=10)


class ModifyFrame(Frame):                      # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master                     # 定义内部变量root
        self.E1 = Entry(self)
        self.E2 = Entry(self)
        self.E3 = Entry(self)
        self.E4 = Entry(self)
        self.E5 = Entry(self)
        self.E6 = Entry(self)
        self.E7 = Entry(self)
        self.E8 = Entry(self)
        self.E9 = Entry(self)
        self.E10 = Entry(self)
        self.E11 = Entry(self)
        self.E12 = Entry(self)
        self.E13 = Entry(self)
        self.E14 = Entry(self)
        self.createPage()

    def Isspace(self, text):                 # 检索text中的内容是否为空，如果是赋值为1
        temp = 0
        for i in text:
            if not i.isspace():
                temp = 1
                break

        if temp == 1:
            return 0
        else:
            return 1

    def modify(self, num, name, major, calss, score1, score2, score3, score4, score5, score6, score7, score8, score9, score10):
        workbook = xlrd.open_workbook("学生信息成绩表.xls")                 # 打开表格
        data = xlrd.open_workbook('学生信息成绩表.xls', formatting_info=True)
        sheet = workbook.sheet_by_index(0)
        my_dict = dict()                                                   # 创建一个字典
        for i in range(2, sheet.nrows):
            my_dict_key = sheet.cell_value(i, 1)
            my_dict_value = sheet.cell_value(i, 2)
            my_dict[my_dict_key] = my_dict_value
        my_dict = list(map(int, my_dict))
        excel = copy(wb=data)                                              # 完成xlrd对象向xlwt对象转换
        excel_table = excel.get_sheet(0)
        dic = {}
        for i in range(2, sheet.nrows):                                   # 循环遍历每一行的内容
            lis = []                                                      # 创建一个新的列表
            for j in range(1, sheet.ncols):
                lis.append(sheet.cell(i, j).value)
            dic[sheet.cell(i, 1).value] = lis
        i = 0
        leap = 0
        a = 0
        for temp in my_dict:                                            # 循环遍历字典里的数据
            a = a + 1
            if temp == int(num):                                        # 判断输入的学号
                leap = 1
                if num != '':                                           # 是否为空
                    dic[num][0] = num                                   # 不是则将输入的学号赋给它
                if name != '':
                    dic[str(num)][1] = name
                if major != '':
                    dic[str(num)][2] = major
                if calss != '':
                    dic[str(num)][3] = calss
                if score1 != '':
                    dic[str(num)][4] = score1
                if score2 != '':
                    dic[str(num)][5] = score2
                if score3 != '':
                    dic[str(num)][6] = score3
                if score4 != '':
                    dic[str(num)][7] = score4
                if score5 != '':
                    dic[str(num)][8] = score5
                if score6 != '':
                    dic[str(num)][9] = score6
                if score7 != '':
                    dic[str(num)][10] = score7
                if score8 != '':
                    dic[str(num)][11] = score8
                if score9 != '':
                    dic[str(num)][12] = score9
                if score10 != '':
                    dic[str(num)][13] = score10
                break
            else:
                i = i + 1
        while leap == 1:
            ncols = 1
            for value in dic[str(num)]:                # 循环遍历字典中的值
                excel_table.write(a+1, ncols, value, set_style('Courier New', 20 * 9, 0x01, 42, 48))    # 将遍历的值依次写入
                ncols = ncols + 1
            messagebox.showinfo(title='提示', message="修改成功")
            break
        else:
            messagebox.showinfo(title='结果', message="不存在该学生信息！")
        excel.save('学生信息成绩表.xls')
        return

    def click(self):
        num = self.E1.get()                      # 得到输入的数据
        name = self.E2.get()
        major = self.E3.get()
        calss = self.E4.get()
        score1 = self.E5.get()
        score2 = self.E6.get()
        score3 = self.E7.get()
        score4 = self.E8.get()
        score5 = self.E9.get()
        score6 = self.E10.get()
        score7 = self.E11.get()
        score8 = self.E12.get()
        score9 = self.E13.get()
        score10 = self.E14.get()
        if self.Isspace(name) or self.Isspace(num) or self.Isspace(major) or self.Isspace(calss):             # 判断是否输入的数据是否为空
            messagebox.showinfo(title='提示', message="输入项为空")
        else:
            self.modify(num, name, major, calss, score1, score2, score3, score4, score5, score6, score7, score8, score9, score10)            # 如果不是就将输入的值进行指定的操作

    def createPage(self):                                            # 设置文本，文本框和按钮组件
        Label(self).grid(row=0, stick=W, pady=10)
        Label(self, text='学号：', font=("华文彩云")).grid(column=0, row=1, stick=W, pady=10)
        self.E1.grid(row=1, column=1, stick=E)
        Label(self, text='姓名：', font=("华文彩云")).grid(column=3, row=1, stick=W, pady=10)
        self.E2.grid(row=1, column=4, stick=E)
        Label(self, text='专业：', font=("华文彩云")).grid(row=3, stick=W, pady=10)
        self.E3.grid(row=3, column=1, stick=E)
        Label(self, text='班级：', font=("华文彩云")).grid(column=3, row=3, stick=W, pady=10)
        self.E4.grid(row=3, column=4, stick=E)
        Label(self, text='高级程序语言：', font=("华文彩云")).grid(row=6, stick=W, pady=10)
        self.E5.grid(row=6, column=1, stick=E)
        Label(self, text='python编程：', font=("华文彩云")).grid(column=3, row=6, stick=W, pady=10)
        self.E6.grid(row=6, column=4, stick=E)
        Label(self, text='数据库原理：', font=("华文彩云")).grid(row=7, stick=W, pady=10)
        self.E7.grid(row=7, column=1, stick=E)
        Label(self, text='数据结构与算法：', font=("华文彩云")).grid(column=3, row=7, stick=W, pady=10)
        self.E8.grid(row=7, column=4, stick=E)
        Label(self, text='数学分析：', font=("华文彩云")).grid(row=9, stick=W, pady=10)
        self.E9.grid(row=9, column=1, stick=E)
        Label(self, text='高等数学：', font=("华文彩云")).grid(column=3, row=9, stick=W, pady=10)
        self.E10.grid(row=9, column=4, stick=E)
        Label(self, text='网络爬虫：', font=("华文彩云")).grid(row=11, stick=W, pady=10)
        self.E11.grid(row=11, column=1, stick=E)
        Label(self, text='数据可视化：', font=("华文彩云")).grid(column=3, row=11, stick=W, pady=10)
        self.E12.grid(row=11, column=4, stick=E)
        Label(self, text='数据挖掘：', font=("华文彩云")).grid(row=13, stick=W, pady=10)
        self.E13.grid(row=13, column=1, stick=E)
        Label(self, text='数据分析：', font=("华文彩云")).grid(column=3, row=13, stick=W, pady=10)
        self.E14.grid(row=13, column=4, stick=E)
        Label(self, text='修改学生信息', font=("华文彩云")).grid(row=0, column=2, stick=N, pady=10)
        Label(self, text='各科成绩', font=("华文彩云")).grid(row=5, column=2, stick=N, pady=10)
        action = ttk.Button(self, text="修改", command=self.click)
        action.grid(row=15, column=4, stick=E, pady=10)


class QueryFrame(Frame):                       # 继承Frame类
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.root = master                     # 定义内部变量root
        self.E1 = Entry(self)
        self.E2 = Entry(self)
        self.boxChoice = Entry(self)
        self.createPage()

    def Isspace(self, text):                  # 同上
        temp = 0
        for i in text:
            if not i.isspace():
                temp = 1
                break

        if temp == 1:
            return 0
        else:
            return 1

    def query(self, num, course):
        workbook = xlrd.open_workbook("学生信息成绩表.xls")
        data = xlrd.open_workbook('学生信息成绩表.xls', formatting_info=True)
        sheet = workbook.sheet_by_index(0)
        my_dict = dict()
        for i in range(2, sheet.nrows):
            my_dict_key = sheet.cell_value(i, 1)
            my_dict_value = sheet.cell_value(i, 2)
            my_dict[my_dict_key] = my_dict_value
        my_dict = list(map(int, my_dict))
        excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
        i = 0
        a = 0
        leap = 0
        for temp in my_dict:                    # 循环遍历字典中的值
            a = a + 1                           # 每次循环a的值都加一，来获取表格的行数
            if temp == int(num):                # 判断数据是否为输入的值
               leap = 1
               if course == "全部":              # 判断是否为全部
                  messagebox.showinfo(title='提示',               # 输出所获取的数据
                                      message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1,
                                                                                                              2) + "\n专业:" + sheet.cell_value(
                                          a + 1, 3) + "\n班级:" + sheet.cell_value(a + 1,
                                                                                 4) + "\n高级程序语言:" + sheet.cell_value(
                                          a + 1, 5) + "\npython编程:" + sheet.cell_value(a + 1,
                                                                                       6) + "\n数据库原理:" + sheet.cell_value(
                                          a + 1, 7) + "\n数据结构与算法:" + sheet.cell_value(a + 1,
                                                                                      8) + "\n数学分析:" + sheet.cell_value(
                                          a + 1, 9) + "\n高等数学:" + sheet.cell_value(a + 1,
                                                                                   10) + "\n网络爬虫:" + sheet.cell_value(
                                          a + 1, 11) + "\n数据可视化:" + sheet.cell_value(a + 1,
                                                                                     12) + "\n数据挖掘:" + sheet.cell_value(a + 1, 13) + "\n数据分析:" + sheet.cell_value(a + 1, 14))
               elif course =="高级程序语言":          # 单独检索成绩
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2)+ "\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n高级程序语言:" + sheet.cell_value(a+1, 5))
               elif course == "python编程":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2)+ "\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\npython编程:" + sheet.cell_value(a+1, 5))
               elif course == "数据库原理":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2)+ "\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n数据库原理:" + sheet.cell_value(a+1, 5))
               elif course == "数据结构与算法":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2) +"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n数据结构与算法:" + sheet.cell_value(a+1, 5))
               elif course == "数学分析":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2) +"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n数学分析:" + sheet.cell_value(a+1, 5))
               elif course == "高等数学":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2)+"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n高等数学:" + sheet.cell_value(a+1, 5))
               elif course == "网络爬虫":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2) +"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n网络爬虫:" + sheet.cell_value(a+1, 5))
               elif course == "数据可视化":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2)+"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n数据可视化:" + sheet.cell_value(a+1, 5))
               elif course == "数据挖掘":
                  messagebox.showinfo(title='提示',message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2) +"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n数据挖掘:" + sheet.cell_value(a+1, 5))
               elif course == "数据分析":
                  messagebox.showinfo(title='提示', message="学号：" + sheet.cell_value(a + 1, 1) + "\n姓名:" + sheet.cell_value(a + 1, 2) +"\n专业:" + sheet.cell_value(a+1, 3) + "\n班级:" + sheet.cell_value(a+1, 4)+"\n数据分析:" + sheet.cell_value(a+1, 5))
                  break
            else:
                i = i +1
        while leap == 0:
          messagebox.showinfo(title='提示', message="不存在该学生信息")
          break
        excel.save('学生信息成绩表.xls')
        return

    def createPage(self):                # 设置文本，文本框和按钮组件
        Label(self).grid(row=0, stick=W, pady=10)
        Label(self, text='查询个人成绩', font=("华文彩云")).grid(row=0, column=1, stick=N, pady=10)
        Label(self, text='学号: ', font=("华文彩云")).grid(row=1, stick=W, pady=10)
        self.E1.grid(row=1, column=1, stick=E)
        Label(self, text='科目: ', font=("华文彩云")).grid(row=2, stick=W, pady=10)
        sexBoxValue = StringVar()
        self.boxChoice = ttk.Combobox(self, width=17, textvariable=sexBoxValue, state='readonly')
        self.boxChoice['value'] = ("高级程序语言", "python编程", "数据库原理", "数据结构与算法", "数学分析", "高等数学", "网络爬虫", "数据可视化", "数据挖掘", "数据分析", "全部")
        self.boxChoice.current(0)
        self.boxChoice.grid(row=2, column=1, sticky=E)
        action = ttk.Button(self, text="查询", command=self.click)
        action.grid(row=6, column=1, stick=E, pady=10)

    def click(self):
        num = self.E1.get()
        course = self.boxChoice.get()
        if self.Isspace(num) or self.Isspace(course):             # 判断是否为空
            messagebox.showinfo(title='提示', message="输入项为空")
        else:
            self.query(num, course)                # 执行操作


class MergeFrame(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.createPage()

    def doProcess(self):                               # 设置表格样式
        workbook1 = xlrd.open_workbook('学生信息表.xlsx')   # 打开Excel表
        workbook2 = xlrd.open_workbook('学生成绩表.xlsx')

        sheet1 = workbook1.sheet_by_name(workbook1.sheet_names()[0])   # 获取工作表
        rows1 = sheet1.nrows  # 获取行
        cols1 = sheet1.ncols  # 获取列

        sheet2 = workbook2.sheet_by_name(workbook2.sheet_names()[0])
        rows2 = sheet2.nrows
        cols2 = sheet2.ncols

        workbook = xlwt.Workbook(encoding="UTF-8")   # 设置新表
        sheets = workbook.add_sheet("Sheet1")

        # 设置列宽
        sheets.col(0).width = 256 * 5
        sheets.col(1).width = 256 * 18
        sheets.col(2).width = 256 * 18
        for i in range(3, 15):
            sheets.col(i).width = 256 * 14

        # 设置行高
        sheets.row(0).height_mismatch = True  # 允许设置行高
        sheets.row(0).height = 20 * 21
        for i in range(1, rows2):
            sheets.row(i).height_mismatch = True
            sheets.row(i).height = 20 * 15

        def set_styles(name, height, borders, colour, colour_index):   # 设置样式
            styles = xlwt.XFStyle()
            styles.font = xlwt.Font()
            styles.font.name = name  # 设置字体
            styles.font.height = height  # 设置字体大小
            styles.font.colour_index = colour_index  # 设置字体颜色
            styles.borders = xlwt.Borders()
            # 设置边框
            styles.borders.left = borders
            styles.borders.top = borders
            styles.borders.right = borders
            styles.borders.bottom = borders
            # 设置背景
            styles.pattern = xlwt.Pattern()
            styles.pattern.pattern = True  # 允许设置背景
            # styles.pattern.pattern_back_colour = 0x04    # 后背景颜色
            styles.pattern.pattern_fore_colour = colour  # 前背景颜色
            # 设置表格对齐属性
            styles.alignment = xlwt.Alignment()
            styles.alignment.horz = xlwt.Alignment.HORZ_CENTER
            styles.alignment.wrap = xlwt.Alignment.NOT_WRAP_AT_RIGHT  # 自动换行
            styles.alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直对齐
            return styles

        sheets.write_merge(0, 0, 0, 13, '学生信息成绩表', set_styles('黑体', 20 * 16, 0x01, 41, 6))  # 合并单元格写入内容
        for i in range(1, rows1):       # 获取行
            for j in range(0, cols1):   # 获取列
                sheets.write(i, j, sheet1.row_values(i)[j], set_styles('Courier New', 20 * 9, 0x01, 42, 48))  # 向指定行列写入信息
        for i in range(1, rows2):
            for j in range(3, cols2):
                sheets.write(i, j + 2, sheet2.row_values(i)[j], set_styles('Courier New', 20 * 9, 0x01, 42, 48))

        workbook.save("学生信息成绩表.xls")  # 保存到新表
        messagebox.showinfo(title='提示', message="合并成功！")

    def createPage(self):       # 设置文本，文本框和按钮组件
        Label(self, text='合并表格', font=("华文彩云")).grid(row=0, column=1, stick=E, pady=10)
        action = ttk.Button(self, text="合并", command=self.doProcess)
        action.grid(row=6, column=1, stick=E, pady=10)


class VisFrame(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.E1 = Entry(self)
        self.createPage()

    def Isspace(self, text):                 # 同上
        temp = 0
        for i in text:
            if not i.isspace():
                temp = 1
                break

        if temp == 1:
            return 0
        else:
            return 1

    def click(self):
        num = self.E1.get()
        if self.Isspace(num) :
            messagebox.showinfo(title='提示', message="输入项为空")
        else:
            matplotlib.rcParams['font.sans-serif'] = ['SimHei']                 # 设置条形图可用汉字和为负数
            matplotlib.rcParams['axes.unicode_minus'] = False

            workbook = xlrd.open_workbook("学生信息成绩表.xls")
            sheet = workbook.sheet_by_index(0)
            my_dict = dict()
            for i in range(2, sheet.nrows):
                my_dict_key = sheet.cell_value(i, 1)
                my_dict_value = sheet.cell_value(i, 2)
                my_dict[my_dict_key] = my_dict_value
            my_dict = list(map(int, my_dict))
            a = 0
            teamp = 0
            list1 = []
            for i in my_dict:
                a = a + 1
                if i == int(num):
                    teamp = 1
                    for j in range(5, sheet.ncols):             # 遍历循环第六行到最后一行的数据
                        list1.append(sheet.cell(a+1, j).value)          # 将遍历的所有数据都放入列表中

                    list1 = list(map(int, list1))
                    N = 10                                      # 定义条形的数目
                    name_list = ['高级程序语言', 'python编程', '数据库原理 ', '数据结构与算法', '数学分析', '高等数学', '网络爬虫', '数据可视化', '数据挖掘', '数据分析']          # 定义x轴的显示数据
                    index = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]              # 定义x轴有10个数据
                    plt.ylim(0, 100)                                # y轴是从0到100
                    ind = np.arange(N)
                    width = 0.45                                    # 间隔宽度

                    p1 = plt.bar(ind, list1, width, facecolor='blue', alpha=0.5, label='分数')           # 设置右上角的标识
                    plt.xticks(index, name_list)
                    plt.xlabel(u"各门学科", fontsize=16)           # 指定x轴描述信息
                    plt.ylabel(u"成绩", fontsize=16)              # 指定y轴描述信息
                    plt.title(sheet.cell_value(a+1, 2)+"的个人成绩条形统计图", fontsize=22)  # 指定图表描述信息
                    plt.hlines(60, -1, 10, linestyles='--')            # 设置再60分时的虚线
                    for rect in p1:                                # 循环遍历将每个的分数显示再数据图的顶部
                      height = rect.get_height()
                      plt.text(rect.get_x() + rect.get_width() / 2, height, height, ha='center', va='bottom', fontsize=14)
                    plt.legend()
                    plt.show()
            if teamp == 1:
                pass
            else:
                messagebox.showinfo(title='提示', message="不存在该学生信息！")

    def createPage(self):                    # 设置文本，文本框和按钮组件
        Label(self).grid(row=0, stick=W, pady=10)
        Label(self, text='导出学生成绩条形图', font=("华文彩云")).grid(row=0, column=1, stick=N, pady=10)
        Label(self, text='学号: ', font=("华文彩云")).grid(row=1, stick=W, pady=10)
        self.E1.grid(row=1, column=1, stick=E)
        action = ttk.Button(self, text="导出", command=self.click)
        action.grid(row=6, column=1, stick=E, pady=10)


class StuFrame(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.E1 = Entry(self)
        self.createPage()

    def Isspace(self, text):                    # 同上
        temp = 0
        for i in text:
            if not i.isspace():
                temp = 1
                break
        if temp == 1:
            return 0
        else:
            return 1

    def click(self):
        num = self.E1.get()
        if self.Isspace(num):
            messagebox.showinfo(title='提示', message="输入项为空")
        else:                    # 设置表格样式
            workbooks = xlwt.Workbook(encoding='UTF-8')
            sheets = workbooks.add_sheet('Sheet1')

            # 设置列宽
            for i in range(0, 8):
                sheets.col(i).width = 256 * 10
            # 设置行高
            sheets.row(0).height_mismatch = True  # 允许设置行高
            sheets.row(0).height = 20 * 46
            for i in range(1, 18):
                sheets.row(i).height_mismatch = True
                sheets.row(i).height = 20 * 20

            def set_sty(name, height, borders, colour, horz):  # 设置样式
                sty = xlwt.XFStyle()
                sty.font = xlwt.Font()
                sty.font.name = name  # 设置字体
                sty.font.height = height  # 设置字体大小
                sty.borders = xlwt.Borders()
                # 设置边框
                sty.borders.left = borders
                sty.borders.top = borders
                sty.borders.right = borders
                sty.borders.bottom = borders
                # 设置背景
                sty.pattern = xlwt.Pattern()
                sty.pattern.pattern = True  # 允许设置背景
                # style.pattern.pattern_back_colour = 0x04    # 后背景颜色
                sty.pattern.pattern_fore_colour = colour  # 前背景颜色
                # 设置表格对齐属性
                sty.alignment = xlwt.Alignment()
                sty.alignment.horz = horz
                sty.alignment.wrap = xlwt.Alignment.NOT_WRAP_AT_RIGHT  # 自动换行
                sty.alignment.vert = xlwt.Alignment.VERT_CENTER  # 垂直对齐
                return sty
            # 合并单元格并写入基本内容，设置样式
            sheets.write_merge(0, 0, 0, 7, '学生个人信息表', set_sty('华文行楷', 20 * 26, 0x07, 45, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(1, 1, 0, 7, '学生个人信息', set_sty('宋体', 20 * 11, 0x01, 51, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(5, 5, 0, 7, '学生成绩信息', set_sty('宋体', 20 * 11, 0x00, 51, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(11, 11, 0, 7, '学生成绩分析', set_sty('宋体', 20 * 11, 0x00, 51, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(15, 15, 0, 7, '学生成绩数据分析', set_sty('宋体', 20 * 11, 0x01, 51, xlwt.Alignment.HORZ_CENTER))
            sheets.write(2, 0, '姓名：', set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
            sheets.write(2, 4, '学号：', set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
            sheets.write(3, 0, '年级：', set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
            sheets.write(3, 4, '班级：', set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
            sheets.write(4, 0, '专业：', set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(6, 6, 0, 1, '高级程序语言', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(7, 7, 0, 1, '数据库原理', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(8, 8, 0, 1, '数学分析', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(9, 9, 0, 1, '网络爬虫', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(10, 10, 0, 1, '数据挖掘', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(6, 6, 4, 5, 'python编程', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(7, 7, 4, 5, '数据结构与算法', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(8, 8, 4, 5, '高等数学', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(9, 9, 4, 5, '数据可视化', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(10, 10, 4, 5, '数据分析', set_sty('宋体', 20 * 11, 0x01, 40, xlwt.Alignment.HORZ_LEFT))
            sheets.write_merge(12, 12, 0, 1, '优科目数', set_sty('宋体', 20 * 11, 0x01, 24, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(13, 13, 0, 1, '中科目数', set_sty('宋体', 20 * 11, 0x01, 24, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(14, 14, 0, 1, '不及格科目数', set_sty('宋体', 20 * 11, 0x01, 24, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(16, 16, 0, 1, '总成绩', set_sty('宋体', 20 * 11, 0x01, 46, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(17, 17, 0, 1, '绩点', set_sty('宋体', 20 * 11, 0x01, 46, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(12, 12, 4, 5, '良科目数', set_sty('宋体', 20 * 11, 0x01, 24, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(13, 13, 4, 5, '及格科目数', set_sty('宋体', 20 * 11, 0x01, 24, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(16, 16, 4, 5, '平均成绩', set_sty('宋体', 20 * 11, 0x01, 46, xlwt.Alignment.HORZ_CENTER))
            sheets.write_merge(17, 17, 4, 5, '总评', set_sty('宋体', 20 * 11, 0x01, 46, xlwt.Alignment.HORZ_CENTER))
            workbooks.save("学生个人信息表.xls")
            workbook = xlrd.open_workbook("学生信息成绩表.xls")
            my_sheet = workbook.sheet_by_index(0)
            # 循环获得sheet的数据，添加至字典中
            my_dict = dict()
            for i in range(2, my_sheet.nrows):             # 循环遍历的方式将数据放入字典中
                my_dict_key = my_sheet.cell_value(i, 1)
                my_dict_value = my_sheet.cell_value(i, 2)
                my_dict[my_dict_key] = my_dict_value
            my_dict = list(map(int, my_dict))             # 将字典数据转化为整形
            tem = -1
            a = 1
            y = 0
            l = 0
            z = 0
            m = 0
            c = 0
            t = 0
            for o in my_dict:         # 循环遍历将数据放入列表
                a = a+1
                if o == int(num):
                    b = []
                    for j in range(5, my_sheet.ncols):
                        b.append(my_sheet.cell(a, j).value)
                    for q in b:                 # 遍历列表中的数据
                        tem = tem + 1
                        if q == '—':           # 如果遇到-就将其转化为0，以免转化为整形时报错
                            b[tem] = 0
                    b = list(map(int, b))
                    p = []
                    for i in b:
                      p.append(int(i))
                    for j in p:                     # 遍历p中的数据判断属于的阶段
                        if (j >= 90) and (j < 100):
                            y = y + 1
                        if (j >= 80) and (j < 90):
                            l = l + 1
                        if (j >= 70) and (j < 80):
                            z = z + 1
                        if (j >= 60) and (j < 70):
                            m = m + 1
                        if j<60:
                            c = c + 1
                # 计算绩点 jd
                    K = []
                    for i in range(0, 10, 2):
                        for j in p:
                            if (j >= 90) and (j < 100):
                                K.append(p[i] * 4)
                                break
                            if (j >= 85) and (j < 90):
                                K.append(p[i] * 3.7)
                                break
                            if (j >= 82) and (j < 85):
                                K.append(p[i] * 3.3)
                                break
                            if (j >= 78) and (j < 82):
                                K.append(p[i] * 3)
                                break
                            if (j >= 75) and (j < 78):
                                K.append(p[i] * 2.7)
                                break
                            if (j >= 71) and (j < 75):
                                K.append(p[i] * 2.3)
                                break
                            if (j >= 66) and (j < 71):
                                K.append(p[i] * 2)
                                break
                            if (j >= 62) and (j < 66):
                                K.append(p[i] * 1.7)
                                break
                            if (j >= 60) and (j < 62):
                                K.append(p[i] * 1.3)
                                break
                            if j < 60:
                                K.append(p[i] * 0)
                                break
                    n = sum(K)
                    jd = n / sum(p)
                    # 计算总成绩 h
                    h = sum(p)
                    # 根据绩点计算优良中差 g
                    g = []
                    if jd >= 4:
                      g = ["优"]
                    if (jd >= 3) and (jd < 4):
                      g = ["良"]
                    if (jd >= 2) and (jd < 3):
                      g = ["中"]
                    if (jd >= 1) and (jd < 2):
                      g = ["及格"]
                    if (jd >= 0) and (jd < 1):
                      g = ["差"]
                    workbook = xlrd.open_workbook("学生个人信息表.xls")
                    data = xlrd.open_workbook("学生个人信息表.xls", formatting_info=True)
                    excel = copy(wb=data)  # 完成xlrd对象向xlwt对象转换
                    excel_table = excel.get_sheet(0)
                    table = data.sheets()[0]
                    # 按位置将所计算出来的数据填入表中，以一定的格式
                    excel_table.write_merge(2, 2, 1, 3, my_sheet.cell_value(a, 2), set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(2, 2, 5, 7, my_sheet.cell_value(a, 1), set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(3, 3, 1, 3, "2019级", set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(3, 3, 5, 7, my_sheet.cell_value(a, 4), set_sty('宋体', 20 * 11, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(4, 4, 1, 7, my_sheet.cell_value(a, 3), set_sty('华文楷体', 20 * 16, 0x01, 31, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(6, 6, 2, 3, my_sheet.cell_value(a, 5), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(6, 6, 6, 7, my_sheet.cell_value(a, 6), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(7, 7, 2, 3, my_sheet.cell_value(a, 7), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(7, 7, 6, 7, my_sheet.cell_value(a, 8), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(8, 8, 2, 3, my_sheet.cell_value(a, 9), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(8, 8, 6, 7, my_sheet.cell_value(a, 10), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(9, 9, 2, 3, my_sheet.cell_value(a, 11), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(9, 9, 6, 7, my_sheet.cell_value(a, 12), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(10, 10, 2, 3, my_sheet.cell_value(a, 13), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(10, 10, 6, 7, my_sheet.cell_value(a, 14), set_sty('宋体', 20 * 11, 0x01, 27, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(12, 12, 2, 3, y, set_sty('宋体', 20 * 11, 0x01, 26, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(12, 12, 6, 7, l, set_sty('宋体', 20 * 11, 0x01, 26, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(13, 13, 2, 3, z, set_sty('宋体', 20 * 11, 0x01, 26, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(13, 13, 6, 7, m, set_sty('宋体', 20 * 11, 0x01, 26, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(14, 14, 2, 3, c, set_sty('宋体', 20 * 11, 0x01, 2, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(16, 16, 2, 3, h, set_sty('宋体', 20 * 11, 0x01, 3, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(16, 16, 6, 7, h / 10.0, set_sty('宋体', 20 * 11, 0x01, 3, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(17, 17, 2, 3, jd, set_sty('宋体', 20 * 11, 0x01, 3, xlwt.Alignment.HORZ_CENTER))
                    excel_table.write_merge(17, 17, 6, 7, g, set_sty('宋体', 20 * 11, 0x01, 3, xlwt.Alignment.HORZ_CENTER))
                    excel.save("学生个人信息表.xls")
                    t = 1
                    break
            if t == 1:
                messagebox.showinfo(title='提示', message="打印成功！")
            else:
                messagebox.showinfo(title='提示', message="不存在该学生信息！")

    def createPage(self):                         # 设置文本，文本框和按钮组件
        Label(self).grid(row=0, stick=W, pady=10)
        Label(self, text='打印个人信息表', font=("华文彩云")).grid(row=0, column=1, stick=N, pady=10)
        Label(self, text='学号: ', font=("华文彩云")).grid(row=3, stick=W, pady=10)
        self.E1.grid(row=3, column=1, stick=E)
        action = ttk.Button(self, text="打印", command=self.click)
        action.grid(row=6, column=1, stick=E, pady=10)


class HelpFrame(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.createPage()

    def click(self):
        messagebox.showinfo(title='功能介绍', message="1.该项目主要可用于固定格式的表格的合并。\n2.并对学生的相关信息实现增，删，改，查的功能。\n3.可以绘制个人成绩的条形统计图。\n4.打印个人成绩表格。")

    def click1(self):
        messagebox.showinfo(title='反馈', message="如果在使用过程中遇到什么问题，或者有什么更好的建议，希望能够积极向我们反馈。\n联系电话：19115505545\nQQ：1370969596")

    def click2(self):
        messagebox.showinfo(title='关于我们', message="项目：学生信息成绩管理系统\n成员：余唯炜，印昌盛，肖春林，唐子怡\n实现语言：python\n工具：PyCharm\n于6月19日完成该项目")

    def click3(self):                          # 按钮链接网页跳转
        webbrowser.open("https://www.cqnu.edu.cn/", new=0)

    def createPage(self):                     # 设置文本，文本框和按钮组件
        Label(self, text='谢谢使用！', font =("华文彩云")).grid(row=1, column=2, stick=N, pady=10)
        action = ttk.Button(self, text="功能介绍", command=self.click)
        action.grid(row=6, column=1, stick=E, pady=10)
        action = ttk.Button(self, text="反馈", command=self.click1)
        action.grid(row=6, column=3, stick=E, pady=10)
        action = ttk.Button(self, text="关于我们", command=self.click2)
        action.grid(row=7, column=1, stick=E, pady=10)
        action = ttk.Button(self, text="学校官网", command=self.click3)
        action.grid(row=7, column=3, stick=E, pady=10)

