# -*- coding: utf-8 -*-
 
"""
PyQt5 tutorial 
 
In this example, we determine the event sender
object.
 
author: py40.com
last edited: 2017年3月

"""
import sys
from PyQt5.QtWidgets import (QMainWindow, QTextEdit, 
    QAction, QFileDialog, QApplication)
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QWidget, QPushButton, QLineEdit,QLabel, 
    QInputDialog, QApplication)
import fairies as fa 
import xlrd
import xlwt
from xlutils.copy import copy
import re
import threading
import time
import os

# TODO 写入的进度


global gMessage

gMessage = ''

stop_word = ['of','the','to','an','in']

def remove_punctuation(line):
    rule = re.compile(r"[^a-zA-Z0-9\u4e00-\u9fa5]")
    line = rule.sub('', line)
    return line

def makeS(a):

    lista = a.capitalize().split()
    prefix = ''
    for i,word in enumerate(lista):
        if i == 0:
            if len(word) == 4:
                prefix = word
            elif len(word) > 2:
                prefix = word[:3]            
        elif i == len(lista) - 1:
            if len(word) == 4:
                prefix = prefix + word.capitalize()
            elif len(word) > 2:
                prefix = prefix + word[:3].capitalize()
        else:
            if word in stop_word:
                pass
            else:
                prefix = prefix + word[0].capitalize()                           

    a = remove_punctuation(prefix)
    return a

class Example(QMainWindow):
    
    def __init__(self):
        super().__init__()
        
        self.initUI()
              
    def initUI(self):      
        
        self.filePath = ''

        self.btn = QPushButton('选择需要翻译的excel路径', self)
        self.btn.setGeometry(20,20,600,40)
        self.btn.clicked.connect(self.showDialog)

        self.textEdit = QLineEdit(self)
        self.textEdit.move(20, 60)
        self.textEdit.setGeometry(20, 90, 600, 40) 

        self.lbl = QLabel("Ubuntu", self)
        self.lbl.setText('等待运行...')
        self.lbl.setGeometry(20,230,600,40)

        self.start_btn = QPushButton('开始翻译', self)
        self.start_btn.setGeometry(20,160,600,40)
        self.start_btn.clicked.connect(self.start_theading)

        self.open_btn = QPushButton('打开文件所在位置', self)
        self.open_btn.setGeometry(220,300,200,40)
        self.open_btn.clicked.connect(self.openFile)
        self.open_btn.setVisible(False)

        self.setGeometry(300, 300, 640, 400)
        self.setWindowTitle('词根翻译工具v1.0')
        self.show()
                
    def showDialog(self):
 
        fname = QFileDialog.getOpenFileName(self, 'Open file', '/home')
 
        if fname[0]:
            self.filePath = fname[0]
            self.textEdit.setText(fname[0])  

    def start_theading(self):

        t1 = threading.Thread(target=self.start_translate)
        t2 = threading.Thread(target=self.showMessage)
        t1.start()
        t2.start()

    def showMessage(self):
        
        while True:
            time.sleep(0.5)
            self.lbl.setText(gMessage)

    def start_translate(self):

        global gMessage

        gMessage = '开始读取excel...'

        data = xlrd.open_workbook(self.filePath)
        table = data.sheet_by_index(0)

        gMessage = '正在处理词根...'

        rowNum = table.nrows
        colNum = table.ncols
        words = []
        for i in range(rowNum):
            k = table.cell(i,0).value
            if len(k) > 4:
                if k not in words:
                    words.append(k)
            if len(k) == 4:
                if k[:2] not in words:
                    words.append(k[:2])
                if k[2:] not in words:
                    words.append(k[2:])
            if len(k) == 3:
                words.append(k)
            if len(k) == 2:
                if k not in words:
                    words.append(k)
            if len(k) == 1:
                if k not in words:
                    words.append(k)

        num = 20

        times = int(len(words)/ num) + 1

        gMessage = '正在勤奋的翻译词根...  预计花费' + str(int(len(words)*1.5)) + '秒'

        # final_res 是翻译词典
        final_res = {}


        for i in range(times):
            
            if i != 0:
                gMessage = '正在勤奋的翻译词根...  已完成' + str(i*num) + '/' + str(len(words))
            if i == 0:
                res = fa.zh_to_en(words[:(i+1)*num])
                for j in range((i+1)*num):
                    final_res[words[j]] = res[j]
            elif i == times -1:
                res = res + fa.zh_to_en(words[(i)*num:])
                for j in range(num*i,len(words)):
                    final_res[words[j]] = res[j]
            else:
                res = res + fa.zh_to_en(words[i*num:(i+1)*num])
                for j in range(num*i,num*(i+1)):
                    final_res[words[j]] = res[j]

        gMessage = '准备写入...'

        excel = copy(data)
        worksheet = excel.get_sheet(0)

        for i in range(rowNum):
            keyword  = table.cell(i,0).value

            isk = keyword.encode('UTF-8')
            if isk.isalpha() or isk.isdigit():
                worksheet.write(i, 1, keyword)
                worksheet.write(i, 3, keyword)

            elif len(keyword) == 4:
                start = keyword[:2]
                end = keyword[2:]
                translate = final_res[start].capitalize() + ' ' + final_res[end].capitalize()
                lista = final_res[start].capitalize().split()

                worksheet.write(i, 1, translate)

                temp = makeS(translate)

                worksheet.write(i, 3, temp) 
                        
            elif len(keyword) == 3 or len(keyword) > 4:
                translate = final_res[keyword].capitalize()
                worksheet.write(i, 1, translate)
                temp = makeS(translate)
                worksheet.write(i, 3, temp) 

            elif len(keyword) == 2:
                translate = final_res[keyword].capitalize()
                worksheet.write(i, 1, translate)
                temp = makeS(translate)
                worksheet.write(i, 3, temp)


        cwd = os.getcwd()
        output_file = os.path.join(cwd,'已翻译的词根.xls')
        excel.save(output_file); #保存至result路径

        gMessage = '写入完成...'
        os.system(r"start " + os.getcwd())
        self.open_btn.setVisible(True)

    def openFile(self):
        os.system(r"start " + os.getcwd())

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())



