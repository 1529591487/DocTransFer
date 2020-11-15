# -*- coding: utf-8 -*-
"""
@author: liuzhiwei

@Date:  2020/9/22
"""

import sys
import os
import pythoncom

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFrame, QFormLayout

from PublicWidget import QLineEdit_Can_Drop, LogTab
from PyQt5.QtCore import QThread, pyqtSignal
from Public import getFiles

from docx import Document
from docxcompose.composer import Composer

from win32com import client as wc


class DocTransFormat(QWidget):
    def __init__(self, parent=None):
        super(DocTransFormat, self).__init__(parent)
        self.thd = None
        self.configPath = '.\\config\\'
        self.__widgetInit()
        self.__signalInit()

    def __widgetInit(self):
        mainLayout = QVBoxLayout(self)

        self.frameInput = QFrame(self)

        self.edit_FilePath = QLineEdit_Can_Drop(self, title='file', configPath=self.configPath + 'path.ini')
        self.pushbutton_toText = QPushButton(self)
        self.pushbutton_toText.setText('转成txt')

        frameLayout = QFormLayout(self.frameInput)
        frameLayout.addRow('文件(夹)路径', self.edit_FilePath)
        self.frameInput.setLayout(frameLayout)

        self.pushbutton_toDocx = QPushButton(self)
        self.pushbutton_toDocx.setText('转成Docx')

        self.pushbutton_toDoc = QPushButton(self)
        self.pushbutton_toDoc.setText('转成Doc')

        self.pushbutton_docToText = QPushButton(self)
        self.pushbutton_docToText.setText('文件夹中doc转成txt')

        self.pushbutton_docxToText = QPushButton(self)
        self.pushbutton_docxToText.setText('文件夹中docx转成txt')

        self.pushbutton_docTodocx = QPushButton(self)
        self.pushbutton_docTodocx.setText('文件夹中doc转成docx')

        self.pushbutton_docxTodoc = QPushButton(self)
        self.pushbutton_docxTodoc.setText('文件夹中docx转成doc')

        self.logTab = LogTab(self, configPath=self.configPath, ifRedirect=True)

        mainLayout.addWidget(self.frameInput)
        mainLayout.addWidget(self.pushbutton_toText)
        mainLayout.addWidget(self.pushbutton_toDocx)
        mainLayout.addWidget(self.pushbutton_toDoc)

        mainLayout.addWidget(self.pushbutton_docToText)
        mainLayout.addWidget(self.pushbutton_docxToText)
        mainLayout.addWidget(self.pushbutton_docTodocx)
        mainLayout.addWidget(self.pushbutton_docxTodoc)

        mainLayout.addWidget(self.logTab)

        self.setLayout(mainLayout)
        self.resize(600, 400)

    def __signalInit(self):
        self.pushbutton_toText.clicked.connect(lambda: self.trans(self.edit_FilePath.text(), 'txt'))
        self.pushbutton_toDocx.clicked.connect(lambda: self.trans(self.edit_FilePath.text(), 'docx'))
        self.pushbutton_toDoc.clicked.connect(lambda: self.trans(self.edit_FilePath.text(), 'doc'))

        self.pushbutton_docToText.clicked.connect(lambda: self.transAll('.doc', 'txt'))
        self.pushbutton_docxToText.clicked.connect(lambda: self.transAll('.docx', 'txt'))
        self.pushbutton_docTodocx.clicked.connect(lambda: self.transAll('.doc', 'docx'))
        self.pushbutton_docxTodoc.clicked.connect(lambda: self.transAll('.docx', 'doc'))

    def transAll(self, oldType, newType):
        if True is os.path.isdir(self.edit_FilePath.text()):
            dirPath = self.edit_FilePath.text()
        elif True is os.path.isfile(self.edit_FilePath.text()):
            dirPath = os.path.dirname(self.edit_FilePath.text())
        else:
            self.logError('不是文件夹，也不是文件')
            return
        fileList = getFiles(dirPath, oldType)
        self.thd = TransThread(fileList, newType)
        self.thd.errorSignal.connect(self.logError)
        self.thd.infoSignal.connect(self.logInfo)
        self.thd.start()

    def trans(self, filePath, newType):
        self.thd = TransThread([filePath], newType)
        self.thd.errorSignal.connect(self.logError)
        self.thd.infoSignal.connect(self.logInfo)
        self.thd.start()

    def logError(self, msg):
        self.logTab.Info(str(msg), color='#ff0000')
        self.logTab.Error(str(msg))

    def logInfo(self, msg):
        self.logTab.Info(str(msg))


# 线程
class TransThread(QThread):
    errorSignal = pyqtSignal(str)
    infoSignal = pyqtSignal(str)

    def __init__(self, filePathList, newType):
        super(TransThread, self).__init__()
        self.filePathList = filePathList
        self.newType = newType

    def run(self):
        self.infoSignal.emit("开始转换")
        try:
            pythoncom.CoInitialize()  # 多线程需要先初始化 http://www.mamicode.com/info-detail-1640383.html
            dealer = wc.gencache.EnsureDispatch('Word.Application')
        except Exception as e:
            self.errorSignal.emit(e)
            return
        for filePath in self.filePathList:
            newType = self.newType
            self.processOneFile(dealer,newType, filePath)

        self.infoSignal.emit("All done！")
        dealer.Quit()

    def processOneFile(self, dealer, newType, filePath):
        self.infoSignal.emit(filePath)
        newFileName = self.getNewFileName(newType, filePath)
        if '' == newFileName:
            return
        try:
            filePath = filePath.replace('/', '\\')
            dealer.Documents.Open(filePath)
            # 文档类型可以参考 https://docs.microsoft.com/zh-cn/office/vba/api/word.wdsaveformat
            if 'txt' == newType:
                fileType = 7
                dealer.ActiveDocument.SaveAs(newFileName, FileFormat=fileType, Encoding=65001)
            elif 'docx' == newType:
                fileType = 12
                dealer.ActiveDocument.SaveAs(newFileName, FileFormat=fileType)
            elif 'doc' == newType:
                fileType = 0
                dealer.ActiveDocument.SaveAs(newFileName, FileFormat=fileType)

            dealer.ActiveDocument.Close()
            self.infoSignal.emit("完成！")
        except Exception as e:
            self.errorSignal.emit(str(e))

    def getNewFileName(self, newType, filePath):
        try:
            dirPath = os.path.dirname(filePath)
            baseName = os.path.basename(filePath)
            fileName = baseName.rsplit('.', 1)[0]
            suffix = baseName.rsplit('.', 1)[1]
            if newType == suffix:
                self.errorSignal.emit('类型相同')
                return ''
            newFileName = '{dir}/{fileName}.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            if os.path.exists(newFileName):
                newFileName = '{dir}/{fileName}_new.{suffix}'.format(dir=dirPath, fileName=fileName, suffix=newType)
            return newFileName
        except Exception as e:
            self.errorSignal.emit(str(e))
            return ''


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = DocTransFormat()
    w.show()
    sys.exit(app.exec_())
