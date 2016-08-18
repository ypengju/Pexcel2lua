#!/usr/bin/python
# -*- coding: utf-8 -*-
import tkFileDialog
import tkMessageBox

import xlrd
import os
import re
from Tkinter import *

reload(sys)
sys.setdefaultencoding('utf-8')

class App(Frame):

    def btn_from_click(self):
        self.filenames = tkFileDialog.askopenfilenames()
        # print self.filenames
        self.mylist.delete(0, END)

        if len(self.filenames):
            self.fromBtn['fg'] = "black"

        for file in self.filenames:
            p, f = os.path.split(file)
            self.mylist.insert(END, f)

    def btn_target_click(self):
        self.dirname = tkFileDialog.askdirectory()
        if os.path.isdir(self.dirname):
            self.targetEntry.delete(0,END)
            self.targetEntry.insert(0, self.dirname)
            self.targetBtn['fg'] = "black"

    #判断是否是制定后缀
    def is_suffix_with(self, filepath, suffix):
        isSuffix = False
        suffixstr = os.path.splitext(filepath)[1]
        # print suffixstr
        if suffixstr == suffix:
            isSuffix = True
        return isSuffix

    #清空lua导出目录
    def clean_lua_dir(self):
        if hasattr(self, "dirname"):
            oldfilelist = os.listdir(self.dirname)
            for filename in oldfilelist:
                filepath = os.path.join(self.dirname, filename)
                if os.path.isfile(filepath):
                    os.remove(filepath)

    def convert_num(self, n):
        try:
            if isinstance(n, float):
                if int(n) == n:
                    return int(n)
                else:
                    return n
            else:
                return n
        except ValueError:
            return n

    def convert_one_file(self, item):
        if self.is_suffix_with(item, ".xlsx") or self.is_suffix_with(item, ".xls"):
            data = xlrd.open_workbook(item)  # 打开xlsx文件
            sheet = data.sheet_by_index(0)  # 获取第一张表
            # col,row
            nrows = sheet.nrows  # 行总数
            ncols = sheet.ncols  # 列总数

            print '[%d,%d] %s' % (nrows, ncols, item)

            p, f = os.path.split(item)
            # print f
            filenum = f[0:4]
            luafile = os.path.join(self.dirname, ("data%s.lua" % filenum))
            # print luafile

            lua_f = file(luafile, 'w')
            lua_f.write("--autogen-begin\n")
            lua_f.write("--Data form [%s]\n" % (f))
            # lua_f.write("--Time [%s]\n\n" % (time.strftime( '%Y-%m-%d %X', time.localtime() )))
            lua_f.write("DataTable = DataTable or {} \n\n")
            lua_f.write("DataTable.Data%s = {\n\t" % filenum)
            lua_f.write("Content = {")

            # 写上每列的说明
            lua_f.write("\n--")
            drow_index = 0

            keys = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r",
                    "s", "t", "u", "v", "w", "x", "y", "z", "aa", "ab", "ac", "ad", "ae", "af", "ag", "ah", "ai",
                    "aj", "ak", "al", \
                    "am", "an", "ao", "ap", "aq", "ar", "as", "at", "au", "av", "aw", "ax", "ay", "az", "ba",
                    "bb", "bc", "bd", "be", "bf", "bg", "bh", "bi", "bj", "bk", "bl", \
                    "bm", "bn", "bo", "bp", "bq", "br", "bs", "bt", "bu", "bv", "bw", "bx", "by", "bz"]

            for col_index in range(0, ncols):
                cell_value = sheet.cell(drow_index, col_index).value
                if str(cell_value) == '':
                    print "data[%d][%d] is null" % (drow_index, col_index)
                else:  # 字符串
                    lua_f.write("%s=%s, " % (keys[col_index], str(cell_value)))

            # 读取表数据并写到lua文件
            for row_index in range(1, nrows):
                lua_f.write("\n\t\t")
#                print (sheet.cell(row_index, 0).value)
                indexvalue = int(sheet.cell(row_index, 0).value)
#                print indexvalue
                lua_f.write("[%d] = {" % indexvalue)
                for col_index in range(0, ncols):
                    cell_value = sheet.cell(row_index, col_index).value
                    cell_value = self.convert_num(cell_value)
                    # print cell_value
                    # print type(cell_value)
                    if isinstance(cell_value, int):
                        lua_f.write("%s = %d, " % (keys[col_index], cell_value))
                    elif isinstance(cell_value, float):
                        lua_f.write("%s = %s, " % (keys[col_index], str(cell_value)))
                    else:
                        lua_f.write("%s = [[%s]], " % (keys[col_index], cell_value))
                lua_f.write("},")

            lua_f.write("\n\t},\n}")
            lua_f.write("\n\n")
            lua_f.write("--autogen-end")
            lua_f.close()
        else:
            tkMessageBox.showerror(u"选错文件了", u"请选中Excel文件再操作")

    #导表
    def excel2lua(self):
        if not hasattr(self, "dirname"):
            tkMessageBox.showerror(u"导出目录为空", u"请选择lua文件的导出路径")
            return
        elif not os.path.isdir(self.dirname):
            tkMessageBox.showerror(u"导出目录错误", u"lua文件的导出路径不是目录")
            return

        if hasattr(self, "filenames") and len(self.filenames) > 0:
            self.clean_lua_dir()
            for item in self.filenames:
                self.convert_one_file(item)
            tkMessageBox.showinfo(u"成功", u"导表成功")
        else:
            tkMessageBox.showerror(u"没有选中Excel文件", u"请选中Excel文件再操作")

    def init_ui(self):

        label = Label(self, text=u'Excel文件：', font=('Arial', 10))
        label.grid(row=0, column=0)

        leftFrame = Frame(self)
        leftFrame['bg'] = "green"
        leftFrame.grid(row=0, column=1)
        # leftFrame.pack(side="left")

        #显示选中的Excel文件
        scrollbar = Scrollbar(leftFrame)
        # scrollbar.grid(row=0, column=3)
        scrollbar.pack(side=RIGHT, fill=Y)
        mylist = Listbox(leftFrame, yscrollcommand=scrollbar.set, width=28)
        self.mylist = mylist
        mylist.pack(side=LEFT, fill=BOTH)
        scrollbar.config(command=mylist.yview)
        #
        fromBtn = Button(self, text=u"选择Excel")
        fromBtn['command'] = self.btn_from_click
        fromBtn['fg'] = "red"
        fromBtn['font'] = "Arial 12"
        fromBtn.grid(row=0, column=3)
        self.fromBtn = fromBtn


        label = Label(self, text=u'Lua目录：', font=('Arial', 10))
        # label.grid_location(100, 100)
        label.grid(row=1, column=0)
        #
        bottomFrame = Frame(self, pady=30)
        bottomFrame.grid(row=1, column=1)
        #
        targetEntry = Entry(bottomFrame)
        targetEntry.pack(side="left")
        targetEntry['width'] = 28
        targetEntry.grid(row=0, column=0)
        self.targetEntry = targetEntry

        #
        targetBtn = Button(self, text=u"选择Lua目录")
        targetBtn['command'] = self.btn_target_click
        targetBtn['font'] = "Arial 12"
        targetBtn['fg'] = "red"
        targetBtn.grid(row=1, column=3)
        self.targetBtn = targetBtn


        tranBtn = Button(self, text=u"开始导表")
        tranBtn['command'] = self.excel2lua
        tranBtn.grid(row=2, column=1)

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.init_ui()


# create the application
root = Tk()
myapp = App(master=root)
width = 600
height = 400
#
# here are method calls to the window manager class
#
myapp.master.title("Excel文件导为Lua表")
myapp.master.maxsize(width, height)
myapp.master.minsize(width, height)

# start the program
myapp.mainloop()
