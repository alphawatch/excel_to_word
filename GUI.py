import tkinter as tk
import tkinter.filedialog
from tkinter import ttk
from win32com import client
import os

top = tk.Tk()


# 选择excel表格
def xz1():
    global xls_file
    xls_file = tkinter.filedialog.askopenfilenames()
    if len(xls_file) != 0:
        string_filename = ""
        for i in range(0, len(xls_file)):
            string_filename += str(xls_file[i]) + "\n"
        msg1.config(text=string_filename)
    else:
        msg1.config(text="您没有选择任何文件", foreground="red")


# 选择word模板
def xz2():
    global doc_file
    doc_file = tkinter.filedialog.askopenfilename()
    if doc_file != '':
        msg2.config(text=doc_file)
    else:
        msg2.config(text="您没有选择任何文件", foreground="red")


# 生成报告
def sc():
    print("start……")
    path = os.getcwd()
    print(path)
    excel = client.Dispatch("Excel.Application")
    excel.Visible = True
    word = client.Dispatch("Word.Application")
    word.Visible = True

    print("共选择>"+str(len(xls_file))+"<份立项表，开始生成可研报告…")
    # 打开所选模板
    doc = word.Documents.Open(doc_file)
    # 定义页眉,Sections[2]代表第三节
    header_count = word.ActiveDocument.Sections.Count
    print("共有>"+str(header_count)+"<个小节，替换最后小节页眉的项目名")
    ye_mei = word.ActiveDocument.Sections[header_count-1].Headers[0]
    # 替换页眉内容
    ye_mei.Range.Find.ClearFormatting()
    ye_mei.Range.Find.Replacement.ClearFormatting()
    ye_mei.Range.Find.Execute("页眉", False, False, False, False, False, True, 1, False, "项目名+可研报告", 1)

    # header_no_change = word.ActiveDocument.Sections[0].Headers[0]
    # 模板中>>正文内容<<替换为立项表中对应内容
    word.Selection.Find.ClearFormatting()
    word.Selection.Find.Replacement.ClearFormatting()
    word.Selection.Find.Execute("{年份}", False, False, False, False, False, True, 1, True, "2018", 2)
    word.Selection.Find.Execute("{月份}", False, False, False, False, False, True, 1, True, "09", 2)
    word.Selection.Find.Execute("{待定}", False, False, False, False, False, True, 1, True, "09", 2)


    # 另存为新的文档，保持模板内容不更改
    doc.SaveAs("D:\linux\PyProject\doc23.docx")

    # excel.Quit()
    # word.Quit()
    print("end！")


top.title('立项表 → 可研报告 (for 北京电信)')
top.geometry("1180x600+300+200")
# 设置ttk控件的风格
style = ttk.Style()
style.configure("BW.TLabel", foreground="black", background="white")
# 定义标签和按钮内容
l1 = ttk.Label(text="1、请选择待处理的Excel表格[可多选]：", style="BW.TLabel")
l2 = ttk.Label(text="2、请选择要使用的word模板：", style="BW.TLabel")
l3 = ttk.Label(text="3、请选择报告生成年月：", style="BW.TLabel")
msg1 = ttk.Label(top, text="请选择……", style="BW.TLabel")
msg2 = ttk.Label(top, text="请选择……", style="BW.TLabel")
btn1 = tk.Button(text="浏览…", command=xz1)
btn2 = tk.Button(text="浏览…", command=xz2)
btn3 = tk.Button(text="开始生成", command=sc)
# 年份下拉选择框体
year = tk.StringVar()
chose_year = ttk.Combobox(top, width=10, textvariable=year)
chose_year['values'] = (2017, 2018, 2019, 2020)
# chose_year.pack()
chose_year.current(1)

# 月份下拉选择框体
month = tk.StringVar()
chose_month = ttk.Combobox(top, width=6, textvariable=month)
chose_month['values'] = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
# chose_month.pack()
chose_month.current(8)

# 对标签和按钮进行布局
# l1.place(x=30, y=30, width=220, height=40)
# l2.place(x=30, y=350, width=220, height=40)
# btn1.place(x=250, y=30, width=60, height=40)
# btn2.place(x=250, y=350, width=60, height=40)
# btn3.place(x=250, y=420, width=150, height=60)
# msg1.place(x=320, y=30, width=800, height=300)
# msg2.place(x=320, y=350, width=800, height=40)
l1.grid(row=0, column=0, sticky=tk.NW, ipadx=20, ipady=10, padx=10, pady=10)
l2.grid(row=1, column=0, sticky=tk.NW, ipadx=20, ipady=10, padx=10, pady=10)
l3.grid(row=2, column=0, sticky=tk.NW, ipadx=20, ipady=10, padx=10, pady=10)
btn1.grid(row=0, column=1, sticky=tk.NW, ipadx=20, ipady=10, padx=10, pady=10)
btn2.grid(row=1, column=1, sticky=tk.NW, ipadx=20, ipady=10, padx=10, pady=10)
btn3.grid(row=3, column=1, sticky=tk.NW, ipadx=20, ipady=10, padx=10, pady=10)
msg1.grid(row=0, column=2, sticky=tk.NW, ipadx=350, ipady=150, pady=10)
msg2.grid(row=1, column=2, sticky=tk.NW, ipadx=350, ipady=15, pady=10)
chose_year.grid(row=2, column=1, sticky=tk.W)
chose_month.grid(row=2, column=2, sticky=tk.W)
tk.mainloop()
