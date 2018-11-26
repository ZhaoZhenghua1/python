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

def GetAllFiles(path):
    ret = []
    if os.path.isfile(path):
        if '$' in path:
            return ret
        extension = os.path.splitext(path)[1]
        if extension == '.docx' or extension == '.doc':
            ret.append(path)
    elif os.path.isdir(path):
        list=os.listdir(path)
        for i in range(0, len(list)):
            ret += (GetAllFiles(path + '\\' + list[i]))
    return ret

filePath = r"D:\试题资料\历史试卷\历史学科已收集原试卷：初中（4479）\初中历史（4479）\中考、会考226"#sys.argv[1]r"D:\ttt"#
print("文件路径：" + filePath)
allfiles = GetAllFiles(filePath)

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

knowledgeFile = r'D:\试题资料\历史识别中易出错词汇.xls'#sys.argv[2]#'历史识别中易出错词汇.xls'#sys.argv[2]'政治识别中易出错词汇（含知识点标准形式、非标准形式、单特征词、歧义型知识点）.xlsx'#sys.argv[2]#
print("知识点词汇路径：" + knowledgeFile)
knowledgeValues = ParseExcel(knowledgeFile)

threadLock = threading.Lock()
def GetAvailableFile():
    threadLock.acquire()
    if len(allfiles) != 0:
        temp = allfiles.pop()
    else:
        temp = '';
    threadLock.release()
    return temp

outputPathEquivalence = os.path.dirname(knowledgeFile) + '\\' + os.path.splitext(os.path.basename(knowledgeFile))[0] + 'Equivalence'
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

def GetWordText(file):
    dll = windll.LoadLibrary('CCSInteroperate.dll')
    jsn = c_wchar_p(dll.FileStandardise(file))
    result = json.loads(jsn.value)
    windll.LoadLibrary('OleAut32.dll').SysFreeString(jsn)
    data = result['Data']
    fileSt = json.loads(data)
    if fileSt['Stauts'] != 1:
        error("File:" + file + " Out:" + str(fileSt));
        return ''
    else:
        return fileSt['Data']['Text']


dll = windll.LoadLibrary('AllSubjectsKnowledgeRec.dll')
dll.InitService('172.16.63.20', '9102')

def PreProcessing(subject, text):
    dll = windll.LoadLibrary('AllSubjectsKnowledgeRec.dll')
    jsn = c_wchar_p(dll.PreProcessingInterface(subject, text))
    result = json.loads(jsn.value)
    windll.LoadLibrary('OleAut32.dll').SysFreeString(jsn)
    return result['Data']

def SplitString(string, regex):
    ret = re.split(regex, string)
    ret = [x for x in ret if len(x.strip()) != 0]#去除空句子
    ret = [re.sub(r"\s+", "", sentence, flags=re.UNICODE) for sentence in ret]#去除句子中的空白字符
    return ret

def SplitSentences(text):
    ret = []
    lines = text.splitlines(False)
    for line in lines:
        segs = SplitString(line, r'。|？|！|，|,')
        ret += segs
    return ret

AllSegments = set()
threadLockSeg = threading.Lock()
def AddToBuffer(seg):
    threadLockSeg.acquire()
    AllSegments.add(seg)
    #print(len(AllSegments))
    threadLockSeg.release()

class TestThread(threading.Thread):
     def __init__(self):
        threading.Thread.__init__(self)

     def recognise(self):
        while True:
            path = GetAvailableFile()
            print("dealing==:" + path)
            if len(path) == 0:
                break
            #try:
            if True:
                text = GetWordText(path)
                text = PreProcessing('H', text)
                if len(text) == 0:
                    continue
                sentences =  SplitSentences(text)
                for sent in sentences:
                    global AllSegments
                    AddToBuffer(sent)
            #except Exception as e:
            #    error(traceback.format_exc())
     def run(self):
        self.recognise()

print('------------------------------------GetAllText--------------------------')
threads=[TestThread(),TestThread(),TestThread(),TestThread(),TestThread()]
for t in threads:
    t.start()

for t in threads:
    t.join()
print('------------------------------------FINISH GetAllText--------------------')
sortedSentences = list(AllSegments)
def SortByLen(array):
    return len(str(array))
sortedSentences.sort(key=SortByLen)

def Segment(sentences):
    sendData = json.loads(r'{"SubjectId": "","SentenceList": [],"FormulaList": [],"ImageList": [],"GetConfigInfoURL": "172.16.63.20:9102"}')
    sendData['SubjectId'] = subject
    sendData['SentenceList'] = sentences
    url = r'http://172.16.63.25:10104/api/KRWebApi/ChWordSegmentation'
    header = {'accept': 'application/json', 'Content-Type': 'application/json'}
    ret = requests.post(url= url, data = json.dumps(sendData).encode('utf-8'), headers = header)
    segresult = json.loads(ret.text)
    if segresult['Status'] != 0:
        error(segresult)
        exit(0)
    return segresult['Data']

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

def GetRightPath(sheetName, knowledge):
    fileName = StandardFileName(knowledge)
    path = outputPathEquivalence + '\\' + sheetName
    if not os.path.exists(path):
        os.makedirs(path)
    while os.path.exists(path + '\\' + fileName):
        fileName += ' -'
    path += '\\' + fileName
    if not os.path.exists(path):
        os.makedirs(path)
    return path

def CreateFile(filename):
    workbook = xlsxwriter.Workbook(filename)
    knowSheet = workbook.add_worksheet('知识点语境')
    red = workbook.add_format({'color': 'red'})
    text_wrap = workbook.add_format({'text_wrap': True})
    bold = workbook.add_format({'bold': True})
 
    knowSheet.set_column('A:A', 90)
    #knowSheet.set_column('B:B', 60)
    knowSheet.write('A1', '知识点语境',bold)
    knowSheet.write('B1', '知识点',bold)
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

sentenceAndSegments = Segment(sortedSentences)

isLiKe = subject in ['B','D','E','F']

def Indexes(sentence, seg):
    ret = []
    index = sentence.find(seg)
    while index != -1:
        ret.append(index)
        index = sentence.find(seg, index + len(seg))
    return ret

for values in knowledgeValues:
    sheetName = values['sheetName']
    header = values['header']
    sheetValues = values['sheetValues']
    for values in sheetValues:
        for knowledge in values:
            if knowledge and knowledge.strip():
                path = GetRightPath(sheetName, knowledge)
                row = 2
                workbook,knowSheet,red,text_wrap = CreateFile(path + '\\' + 'output.xls') 
                for sentAndSeg in sentenceAndSegments:
                    SentenceContent = sentAndSeg['SentenceContent']
                    if len(SentenceContent.strip()) == 0:
                        continue
                    SentenceSegWords = sentAndSeg['SentenceSegWords']
                    string_parts = []
                    if isLiKe:#添加理科处理逻辑
                        if knowledge not in SentenceSegWords:#理科
                            continue
                        for seg in SentenceSegWords:
                            if seg == knowledge:
                                string_parts.append(red)
                            string_parts.append(seg)
                            string_parts.append('\\')
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
                workbook.close()        
                os.rename(path + '\\' + 'output.xls', path + '\\' + 'output' + str(row) + '.xls')

    

