import os
import tkinter
from tkinter.constants import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

tk = tkinter.Tk()
subjects = ['数学','物理','化学','生物']
ids = ['B','D','E','F']
def SubjectID(subject):
    if subject not in subjects:
        messagebox.showinfo("学科错误", "subject not in " + str(subjects))
    return ids[subjects.index(subject)]
SVdataFile = tkinter.StringVar()
def SelectDataFile():
    fileSelected = filedialog.askopenfile(mode='r', filetypes=[('Text','*.txt')]).name
    SVdataFile.set(fileSelected)

SVExcelFile = tkinter.StringVar()
def SelectKnowExcelFile():
    myFormats = [('Excel','*.xls;*.xlsx')]
    fileSelected = filedialog.askopenfile(mode='r', filetypes=myFormats).name
    SVExcelFile.set(fileSelected)

frame = tkinter.Frame(tk, relief=RIDGE, borderwidth=2)
frame.pack(fill=BOTH,expand=1)

button = tkinter.Button(frame, text="选择数据文件", command = SelectDataFile)
button.grid(row=0, column=0, columnspan=1)
#button.pack( side = 'left', anchor= 'n')#,

textFolder = tkinter.Entry(frame, textvariable=SVdataFile, width = 50)
textFolder.grid(row=0, column=1, columnspan=1)

labelSubject = tkinter.Label(frame, text="学科：")
labelSubject.grid(row = 1, column = 0, columnspan = 1)

comboSubject = ttk.Combobox(frame, width = 47, values = subjects)
comboSubject.grid(row=1, column=1, columnspan=1)

label = tkinter.Button(frame, text="选择知识点列表文件：",command = SelectKnowExcelFile)
label.grid(row=2, column = 0, columnspan = 1)

textOutput = tkinter.Entry(frame, textvariable=SVExcelFile, width = 50)
textOutput.grid(row=2, column=1, columnspan=1)

SVOutFolder = tkinter.StringVar()
def SelectOutFolder():
    fileSelected = filedialog.askdirectory()
    SVOutFolder.set(fileSelected)
btnOutput = tkinter.Button(frame, text="输出文件夹：",command = SelectOutFolder)
btnOutput.grid(row=3, column = 0, columnspan = 1)

textOutput = tkinter.Entry(frame, textvariable=SVOutFolder, width = 50)
textOutput.grid(row=3, column=1, columnspan=1)

def ExportKnowledgeItems():
    dataFile = SVdataFile.get()
    subject = SubjectID(comboSubject.get())
    excelFile = SVExcelFile.get()
    outFolder = SVOutFolder.get()
    tk.destroy()
    os.system('python.exe 抓取试题分词知识点导出.py '+ dataFile + ' ' + subject + ' ' + excelFile + ' ' + outFolder)
btnRec = tkinter.Button(frame,text="导出知识点", command = ExportKnowledgeItems)
btnRec.grid(row=4, column=1, columnspan=1)

tk.mainloop()
