import os
import tkinter
from tkinter.constants import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from 化学有机物匹配规则 import ChemistryOrganicMatch

tk = tkinter.Tk()

files = []
iCurFile = 0
SVdataFile = tkinter.StringVar()

def ShowFile():
    global iCurFile
    if iCurFile >= len(files):
        return
    textOutput.delete('1.0', END)
    from 输出所有试题文本到文件中 import GetWordText
    textOutput.insert(INSERT, GetWordText(files[iCurFile]))
    iCurFile += 1
    pass
def SelectFiles():
    fileSelected = filedialog.askopenfiles(mode='r', filetypes=[('Text','*.txt;*.doc;*.docx')])
    print(fileSelected)
    global files
    files = [];
    global iCurFile
    iCurFile=0
    for sel in fileSelected:
        files.append(sel.name)
    SVdataFile.set(str(files))
    ShowFile()

def SelectFolder():
    folderSelected = filedialog.askdirectory()
    print(folderSelected)
    global files
    from 输出所有试题文本到文件中 import GetAllFiles
    files = GetAllFiles(folderSelected);iCurFile=0
    SVdataFile.set(folderSelected)
    ShowFile()

SVExcelFile = tkinter.StringVar()
def SelectKnowExcelFile():
    myFormats = [('Excel','*.xls;*.xlsx')]
    fileSelected = filedialog.askopenfile(mode='r', filetypes=myFormats).name
    SVExcelFile.set(fileSelected)

frame = tkinter.Frame(tk, relief=RIDGE, borderwidth=2)
frame.pack(side=TOP,fill=BOTH, expand = True)

frame0 = tkinter.Frame(frame, relief=RIDGE, borderwidth=2)
frame0.pack(side=TOP,fill=X)
frame1 = tkinter.Frame(frame, relief=RIDGE, borderwidth=2)
frame1.pack(side=TOP,fill=BOTH, expand = True)
frame2 = tkinter.Frame(frame, relief=RIDGE, borderwidth=2)
frame2.pack(side=TOP,fill=X)

button = tkinter.Button(frame0, text="选择文件", command = SelectFiles)
button.pack(padx=5, pady=5, side=LEFT)

button = tkinter.Button(frame0, text="选择文件夹", command = SelectFolder)
button.pack(padx=5, pady=5, side=LEFT)

textFolder = tkinter.Entry(frame0, textvariable=SVdataFile)
textFolder.pack(padx=5, pady=5, side=LEFT,fill=X, expand = True)

button = tkinter.Button(frame0, text="Next==>>", command = ShowFile)
button.pack(padx=5, pady=5, side=LEFT)

class CustomText(tkinter.Text):
    def __init__(self, *args, **kwargs):
        tkinter.Text.__init__(self, *args, **kwargs)
        self.bind_all('<<Modified>>', self._beenModified)
    
    tt = 0
    def _beenModified(self, event=None):
        '''
        Call the user callback. Clear the Tk 'modified' variable of the Text.
        '''
        self.tt += 1
        if self.tt % 2 == 0:
            return
        # If this is being called recursively as a result of the call to
        # clearModifiedFlag() immediately below, then we do nothing.
        #if self._resetting_modified_flag: return

        # Clear the Tk 'modified' variable.
        self.clearModifiedFlag()

        # Call the user-defined callback.
        self.beenModified(event)
    def beenModified(self, event=None):
        '''
        Override this method in your class to do what you want when the Text
        is modified.
        '''
        for iline in range(1, int(self.index('end').split('.')[0])):
            line = self.get("%d.0"%iline,"%d.99999999"%iline)
            if len(line) == 0:
                continue
            
            orgmatch = ChemistryOrganicMatch(line)
            for match in orgmatch:
                self.tag_add("organic", "%d.%d"%(iline,match.mIndex), "%d.%d"%(iline,match.mIndex+match.mLength))

    def clearModifiedFlag(self):
        '''
        Clear the Tk 'modified' variable of the Text.

        Uses the _resetting_modified_flag attribute as a sentinel against
        triggering _beenModified() recursively when setting 'modified' to 0.
        '''
        self.tag_delete('organic')
        self.tag_config("organic", background="yellow", foreground="red")
        # Set the sentinel.
        self._resetting_modified_flag = True

        try:

            # Set 'modified' to 0.  This will also trigger the <<Modified>>
            # virtual event which is why we need the sentinel.
            self.tk.call(self._w, 'edit', 'modified', 0)

        finally:
            # Clean the sentinel.
            self._resetting_modified_flag = False
    

   
textOutput = CustomText(frame1)
textOutput.pack(padx=5, pady=5, fill=BOTH, side=LEFT, expand = True)

textOutput.insert(INSERT, "Hello....3-羟基苯基氧磷基丙酸.")


import xlsxwriter
def export():
    from 输出所有试题文本到文件中 import GetWordText
    from 输出所有试题文本到文件中 import PreProcessing
    from 输出所有试题文本到文件中 import SplitSentences
    fileSelected = filedialog.asksaveasfilename(filetypes=[('xls','*.xls;*.xlsx')])
    if len(fileSelected) == 0:
        return
    filename, file_extension = os.path.splitext(fileSelected)
    if file_extension == "":
        fileSelected += '.xls'
    workbook = xlsxwriter.Workbook(fileSelected)
    knowSheet = workbook.add_worksheet('知识点语境')
    red = workbook.add_format({'color': 'red'})
    text_wrap = workbook.add_format({'text_wrap': True})
    bold = workbook.add_format({'bold': True})
 
    knowSheet.set_column('A:A', 180)
    AllSegments = set()
    for path in files:
        print('reading ' + path)
        text = GetWordText(path)
        text = PreProcessing('化学', text)
        if len(text) == 0:
            continue
        sentences = SplitSentences(text)
        for sent in sentences:
            AllSegments.add(sent)
    print('matching...')
    sortedSentences = list(AllSegments)
    def SortByLen(array):
        return len(str(array))
    sortedSentences.sort(key=SortByLen)
    row = 0
    for sent in sortedSentences:
        orgmatch = ChemistryOrganicMatch(sent)  
        if len(orgmatch) != 0:     
            print( '---row---' + str(row))
            string_parts = [] 
            start = 0
            for item in orgmatch:
                if item.mIndex != start:
                    string_parts.append(orgmatch.mText[start:item.mIndex])
                start = item.mIndex + item.mLength
                string_parts.append(red)
                string_parts.append(str(item))
                print(str(item))
            if start != len(sent):
                string_parts.append(sent[start:])
            if len(string_parts) == 2:
                string_parts += ' '
            string_parts.append(text_wrap)
            knowSheet.write_rich_string('A' + str(row),*string_parts)
            row += 1
            pass
    if row > 0:
        print('saving ' + fileSelected)
        workbook.close() 
        print('done.')
        import subprocess
        fileSelected = fileSelected.replace(r'/', '\\')
        subprocess.Popen(r'explorer /e,/select,"' + fileSelected + r'"')

btnRec = tkinter.Button(frame2,text="导出", command = export)
btnRec.pack(padx=5, pady=5)

tk.mainloop()