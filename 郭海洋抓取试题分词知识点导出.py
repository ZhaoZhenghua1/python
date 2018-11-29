import os
import json
import sys, getopt
import threading
import io
from ctypes import *
import re
import _thread
import shutil
from xlrd import open_workbook
import xlwt
import xlsxwriter
import requests

filePath = sys.argv[1]
print("文件路径：" + filePath)

def ParseExcel(file):
    wb = open_workbook(file)
    lst = []
    for s in wb.sheets():
        values = {}
        sheetName = s.name
        print(sheetName)
        if s.nrows == 0:
            continue
        colNum = s.ncols
        header = []
        for col in range(0, colNum):
            header.append(s.row(0)[col].value)
        sheetValues = []
        for row in range(1, s.nrows):
            vals = []
            for col in range(0, colNum):
                #print(s.row(row)[col].value)
                vals.append(s.row(row)[col].value)
            sheetValues.append(vals)
            
        values['sheetName'] = sheetName  
        values['header'] = header  
        values['sheetValues'] = sheetValues  
        lst.append(values)
    return lst

knowledgeFile = sys.argv[2]#
print("知识点词汇路径：" + knowledgeFile)
knowledgeValues = ParseExcel(knowledgeFile)

#等价关系路径
outputPathEquivalence = os.path.dirname(knowledgeFile) + '\\' + os.path.splitext(os.path.basename(knowledgeFile))[0] + 'Equivalence'
#包含关系路径
outputPathInclusion = os.path.dirname(knowledgeFile) + '\\' + os.path.splitext(os.path.basename(knowledgeFile))[0] + 'Inclusion'
import shutil
print('output Equivalence:' + outputPathEquivalence)
print('output Inclusion:' + outputPathInclusion)
shutil.rmtree(outputPathEquivalence, ignore_errors=True)
shutil.rmtree(outputPathInclusion, ignore_errors=True)

if not os.path.exists(outputPathEquivalence):
    os.makedirs(outputPathEquivalence)
if not os.path.exists(outputPathInclusion):
    os.makedirs(outputPathInclusion)
document = open(outputPathEquivalence + '\\' + os.path.splitext(os.path.basename(knowledgeFile))[0] + " error.txt","wb")

threadLockError = threading.Lock()
def error(str):
    threadLockError.acquire()
    print('PRINT:' + str)
    document.write((str + '\r\n').encode('utf-16'))
    threadLockError.release()

def StandardFileName(file):
    file = file.replace('\\','＼')
    file = file.replace('/','／')
    file = file.replace(':','：')
    file = file.replace('*','＊')
    file = file.replace('?','？')
    file = file.replace('"','＂')
    file = file.replace('>','＞')
    file = file.replace('<','＜')
    file = file.replace('|','｜')
    return file

pathlock = threading.Lock()
def doPath(path, sheetName, fileName):
    pathlock.acquire()
    path = path + '\\' + sheetName
    if not os.path.exists(path):
        os.makedirs(path)
    while os.path.exists(path + '\\' + fileName):
        fileName += ' -'
    path += '\\' + fileName
    if not os.path.exists(path):
        os.makedirs(path)
    pathlock.release()
    return path
def GetERightPath(sheetName, knowledge):
    fileName = StandardFileName(knowledge)
    return doPath(outputPathEquivalence, sheetName, fileName)
def GetIRightPath(sheetName, knowledge):
    fileName = StandardFileName(knowledge)
    return doPath(outputPathInclusion, sheetName, fileName)

def CreateFile(filename, type):
    workbook = xlsxwriter.Workbook(filename)
    knowSheet = workbook.add_worksheet('知识点语境')
    red = workbook.add_format({'color': 'red'})
    text_wrap = workbook.add_format({'text_wrap': True})
    bold = workbook.add_format({'bold': True})
 
    knowSheet.set_column('A:A', 90)
    #knowSheet.set_column('B:B', 60)
    knowSheet.write('A1', '知识点语境',bold)
    if type == 0:
        knowSheet.write('B1', '知识点',bold)
    elif type == 1:
        knowSheet.write('B1', '知识点分词',bold)
        knowSheet.write('C1', '知识点',bold)
    knowSheet.freeze_panes(1, 0)
    
    return workbook,knowSheet,red,text_wrap

def GetSubject(path):
    subjects={'数学':'B','物理':'D','化学':'E','生物':'F','政治':'G','历史':'H','地理':'I', '语文':'A', '英语':'C'}
    for key in subjects:
        if path.find(key) != -1:
            return subjects[key]
    return ""
subject = GetSubject(knowledgeFile)
print('subject:' + subject)

def GetReadData():
    arr = []
    read = open(filePath + r'\allfiledata.txt', 'r',encoding = 'utf-16')
    line = read.readline()
    while line:
        if line.startswith(u'\ufeff'):
            line = line.encode('utf8')[3:].decode('utf8')
        arr.append(json.loads(line))
        line = read.readline()
    read.close()
    return arr
sentenceAndSegments = GetReadData()

isLiKe = subject in ['B','D','E','F']

def GetAvailableKnowledge():
    for values in knowledgeValues:
        sheetName = values['sheetName']
        header = values['header']
        sheetValues = values['sheetValues']
        for values in sheetValues:
            for knowledge in values:
                yield sheetName, knowledge

def LongString(length, string):
    string = '0'*(length - len(string)) + string
    return string
class SortByKnowledge:
    def __init__(self, knowledge):
        self.knowledge = knowledge
    def __call__(self, array):
        knowsegs = []
        for seg in array:
            if seg.find(self.knowledge) != -1 and self.knowledge != seg:
                if seg not in knowsegs:
                    knowsegs.append(seg)
        return LongString(3,str(len(knowsegs))) + LongString(50, knowsegs[0]) + LongString(3, str(len(str(array)))) + LongString(500, str(array))

def DoInclude(IncludeArraySegs, sheetName, knowledge):
    path = GetIRightPath(sheetName, knowledge)
    IncludeArraySegs.sort(key=SortByKnowledge(knowledge))
    allCount = 0
    row = 2
    currentSeg = ''
    filename = StandardFileName(knowledge)
    workbook,knowSheet,red,text_wrap = CreateFile(path + '\\' + filename + '.xls', 1)
    for segs in IncludeArraySegs:
        string_parts = []
        knowSegs = []
        for seg in segs:
            index = seg.find(knowledge) 
            if index == -1:
                string_parts.append(seg)
            else:
                if seg not in knowSegs:
                    knowSegs.append(seg)
                start = seg[0:index]
                if len(start) != 0:
                    string_parts.append(start)
                if knowSegs[0] != currentSeg:
                    currentSeg = knowSegs[0]
                    if row > 1001:
                        workbook.close()        
                        os.rename(path + '\\' + knowledge + '.xls', path + '\\' + knowledge + ' ' + str(allCount) + '.xls')
                        workbook,knowSheet,red,text_wrap = CreateFile(path + '\\' + knowledge + '.xls', 1)
                        row = 2
                string_parts.append(red)
                string_parts.append(knowledge)
                end = seg[index + len(knowledge) : len(seg)]
                if len(end) != 0:
                    string_parts.append(end)
            string_parts.append('\\')
        string_parts = string_parts[:-1]
        if len(string_parts) == 2:
            string_parts += ' '
        string_parts.append(text_wrap)
        knowSheet.write_rich_string('A' + str(row),*string_parts)
        knowSheet.write('B' + str(row), "|".join(knowSegs))
        knowSheet.write('C' + str(row), knowledge)
        row += 1
        allCount += 1
    if row > 2:
        workbook.close()        
        os.rename(path + '\\' + filename + '.xls', path + '\\' + filename + ' ' + str(allCount) + '.xls')
    temp = ((' ' + str(allCount)) if allCount > 0 else ' Empty')
    while os.path.exists(path + temp):
        temp += ' -'
    os.rename(path, path + temp)

availableKnowIte = GetAvailableKnowledge()
threadLockSeg = threading.Lock()
def GetKnow():
    threadLockSeg.acquire()
    return next(availableKnowIte)
    threadLockSeg.release()

class DThread(threading.Thread):
     def __init__(self):
        threading.Thread.__init__(self)
        self.lock = threading.Lock()

     def run(self):
        while True:
            self.lock.acquire()
            try:
                sheetName, knowledge = next(availableKnowIte)
            except StopIteration:
                break
            self.lock.release()
            if knowledge and knowledge.strip():
                filename = StandardFileName(knowledge)
                print("Doing--:" + knowledge)
                path = GetERightPath(sheetName, knowledge)
                IncludeArray = []
                allCount = 0
                row = 2
                workbook,knowSheet,red,text_wrap = CreateFile(path + '\\' + knowledge + '.xls', 0) 
                for sentAndSeg in sentenceAndSegments:
                    if row > 1001:
                        workbook.close()        
                        os.rename(path + '\\' + filename + '.xls', path + '\\' + filename + ' ' + str(allCount) + '.xls')
                        workbook,knowSheet,red,text_wrap = CreateFile(path + '\\' + filename + '.xls', 1)
                        row = 2
                    SentenceContent = sentAndSeg['s']
                    if len(SentenceContent.strip()) == 0:
                        continue
                    SentenceSegWords = sentAndSeg['g']
                    string_parts = []
                    if isLiKe:#添加理科处理逻辑
                        #包含关系
                        if str(SentenceSegWords).find(knowledge)!= -1 and knowledge not in SentenceSegWords:
                            IncludeArray.append(SentenceSegWords)
                        #处理等价关系
                        if knowledge in SentenceSegWords:#理科
                            for seg in SentenceSegWords:
                                if seg == knowledge:
                                    string_parts.append(red)
                                string_parts.append(seg)
                                string_parts.append('\\')
                        else:
                            continue
                    else:#文科              
                        index = SentenceContent.find(knowledge)        
                        if index == -1:
                            continue   

                        begin = 0
                        for seg in SentenceSegWords:
                            end = begin + len(seg)
                            left = max(index, begin)
                            right = min(index + len(knowledge), end)
                            if right > left:
                                string_parts.append(red)
                                string_parts.append(SentenceContent[left:right])
                                if end > right:
                                    string_parts.append(SentenceContent[right:end])
                            else:
                                string_parts.append(seg)
                            if end >= index + len(knowledge) and index != -1:
                                index = SentenceContent.find(knowledge, end)
                            string_parts.append('\\')
                            begin += len(seg)
                    string_parts = string_parts[:-1]
                    if len(string_parts) == 2:
                        string_parts += ' '
                    string_parts.append(text_wrap)
                    knowSheet.write_rich_string('A' + str(row),*string_parts)
                    knowSheet.write('B' + str(row), knowledge)
                    row += 1
                    allCount += 1
                if row > 2:
                    workbook.close()        
                    os.rename(path + '\\' + filename + '.xls', path + '\\' + filename + ' ' + str(allCount) + '.xls')
                temp = ((' ' + str(allCount)) if allCount > 0 else ' Empty')
                while os.path.exists(path + temp):
                    temp += ' -'
                os.rename(path, path + temp)
                DoInclude(IncludeArray, sheetName, knowledge)


print('------------------------------------ParseKnowledge--------------------------')
threadNum = 10
threads = []
for i in range(0, threadNum):
    threads.append(DThread())
for t in threads:
    t.start()
for t in threads:
    t.join()
print('------------------------------------FINISH ParseKnowledge--------------------')
