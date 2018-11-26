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

filePath = r"D:\历史ttt"#sys.argv[1]r"D:\ttt"#
print("文件路径：" + filePath)
allfiles = GetAllFiles(filePath)

threadLock = threading.Lock()
def GetAvailableFile():
    threadLock.acquire()
    if len(allfiles) != 0:
        temp = allfiles.pop()
    else:
        temp = '';
    threadLock.release()
    return temp

document = open(filePath + '\\' + " error.txt","wb")

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
    arrRet = []
    begin = 0
    while begin < len(sentences):
        arrSend = sentences[begin : begin + 10000]
        begin += 10000
        sendData = json.loads(r'{"SubjectId": "","SentenceList": [],"FormulaList": [],"ImageList": [],"GetConfigInfoURL": "172.16.63.20:9102"}')
        sendData['SubjectId'] = subject
        sendData['SentenceList'] = arrSend
        url = r'http://172.16.63.25:10104/api/KRWebApi/ChWordSegmentation'
        header = {'accept': 'application/json', 'Content-Type': 'application/json'}
        ret = requests.post(url= url, data = json.dumps(sendData).encode('utf-8'), headers = header)
        segresult = json.loads(ret.text)
        if segresult['Status'] != 0:
            error(segresult)
            exit(0)
        arrRet += segresult['Data']
    return arrRet

def GetSubject(path):
    subjects={'数学':'B','物理':'D','化学':'E','生物':'F','政治':'G','历史':'H','地理':'I', '语文':'A', '英语':'C'}
    for key in subjects:
        if path.find(key) != -1:
            return subjects[key]
    return ""
subject = GetSubject(filePath)
print('subject:' + subject)

sentenceAndSegments = Segment(sortedSentences)

write = open(filePath + r'\allfiledata.txt', "wb")
write.write(json.dumps(sentenceAndSegments,ensure_ascii=False).encode('utf-16'))
write.close()
#import pickle
#with open(filePath + r'\allfiledata.text', 'wb') as fp:
#    pickle.dump(sentenceAndSegments, fp)