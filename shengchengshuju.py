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
v = tkinter.StringVar()
def SelectFolder():
    folder_selected = filedialog.askdirectory()
    v.set(folder_selected)
def CreateRecData():
    folder = textFolder.get()
    outputfile = textOutput.get()
    subject = SubjectID(comboSubject.get())
    tk.destroy()
    os.system('python.exe 输出所有试题文本到文件中.py '+ folder + ' ' + subject + ' ' + outputfile)
frame = tkinter.Frame(tk, relief=RIDGE, borderwidth=2)
frame.pack(fill=BOTH,expand=1)

button = tkinter.Button(frame, text="选择试题文件夹", command = SelectFolder)
button.grid(row=0, column=0, columnspan=1)
#button.pack( side = 'left', anchor= 'n')#,

textFolder = tkinter.Entry(frame, textvariable=v, width = 50)
textFolder.grid(row=0, column=1, columnspan=1)

labelSubject = tkinter.Label(frame, text="学科：")
labelSubject.grid(row = 1, column = 0, columnspan = 1)

comboSubject = ttk.Combobox(frame, width = 47, values = subjects)
comboSubject.grid(row=1, column=1, columnspan=1)

SVfileSelected = tkinter.StringVar()
def SeleceSavedFile():
    options = {}
    options['defaultextension'] = ".txt"
    options['filetypes'] = [('text files', '.txt')]
    options['initialdir'] = None
    options['initialfile'] = None
    options['title'] = None
    fileselected = filedialog.asksaveasfile(mode='w', **options).name
    SVfileSelected.set(fileselected)
label = tkinter.Button(frame, text="试题数据保存为：", command = SeleceSavedFile)
label.grid(row=2, column = 0, columnspan = 1)

textOutput = tkinter.Entry(frame, width = 50, textvariable = SVfileSelected)
textOutput.grid(row=2, column=1, columnspan=1)
#textFolder.pack(side = 'left')#,fill=BOTH( side = 'left', anchor = 'w', expand=1)#anchor= 'n',

btnRec = tkinter.Button(frame,text="生成试题数据", command = CreateRecData)
btnRec.grid(row=3, column=1, columnspan=1)
#btnRec.pack(side='bottom')

tk.mainloop()