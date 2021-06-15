# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'loader.ui'
#
# Created by: PyQt5 UI code generator 5.12.3
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
import MailMan

class Ui_Form(object): #ui - mailman front page dimension and looks
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(511, 174)
        Form.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(250, 239, 142), stop:1 rgba(255, 255, 255, 255));")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(130, 33, 360, 21))
        font = QtGui.QFont()
        font.setFamily("Century Gothic")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setStyleSheet("background-color: rgb(0,0,0,0%);\n"
"color: rgb(0, 109, 163);")
        self.label_2.setObjectName("label_2")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(10, 40, 101, 120))
        self.label.setStyleSheet("background-color: rgb(0,0,0,0%);")
        self.label.setText("")
        self.label.setPixmap(QtGui.QPixmap("images/mail.png"))
        self.label.setObjectName("label")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(130, 90, 355, 23))
        self.progressBar.setStyleSheet("background-color: rgb(170, 85, 0);")
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(129, 130, 321, 20))
        font = QtGui.QFont()
        font.setFamily("Sitka Text")
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setStyleSheet("background-color: rgb(0,0,0,0%);\n"
"color: rgb(0, 109, 163);")
        self.label_3.setObjectName("label_3")

        self.retranslateUi(Form)
        self.form_data = Form
        QtCore.QMetaObject.connectSlotsByName(Form)
        self.start_loading()

    def start_loading(self): #loads the py module, it start in main event loop
        self.label_3.setText('--->>>>>> Initiating Process... <<<<<<<---')
        self.progressBar.setValue(2)
        self.worker = WorkerThread(None)
        self.worker.start()#
        self.worker.update_progress.connect(self.update_label)
        self.worker.update_bar.connect(self.update_prog_bar)
        #time.sleep(3)
        self.worker.loading.connect(self.gather_module)
        self.worker.finished.connect(self.evt_finished)

    def gather_module(self,val): #gather all modules used in all files
        if val[0]=='opx':
            self.opx = val[1]
        if val[0]=='sys':
            self.sys = val[1]
        if val[0]=='os':
            self.os = val[1]
        if val[0]=='wi':
            self.wi = val[1]
        if val[0]=='wc':
            self.wc = val[1]
        if val[0]=='zip_longest':
            self.zip_longest = val[1]
        if val[0]=='re':
            self.re = val[1]
        if val[0]=='time':
            self.time = val[1]
        if val[0]=='json':
            self.json = val[1]
        if val[0]=='datetime':
            self.datetime = val[1]
        if val[0]=='urlparse':
            self.urlparse = val[1]
        if val[0]=='glob':
            self.glob = val[1]
        if val[0]=='getpass':
            self.getpass = val[1]
        if val[0]=='Path':
            self.Path = val[1]
        if val[0]=='glob':
            self.glob = val[1]

    def evt_finished(self): #mail-man->mail.py, loading ui is finished
        self.label_3.setText('Loading Almost finished ..')
        self.cnf_window = QtWidgets.QMainWindow() #show window ui
        self.cnf_ui = MailMan.Ui_MainWindow(self.getpass,self.datetime,self.wc,self.Path,self.wi,self.os,self.opx,self.sys,self.json,self.glob,self.urlparse,self.time,self.re)
        self.cnf_ui.setupUi(self.cnf_window)
        self.cnf_window.show()
        self.form_data.setVisible(False)

    def update_label(self,val):
        self.label_3.setText(str.strip(val))

    def update_prog_bar(self,val):
        self.progressBar.setValue(int(val))

    def retranslateUi(self, Form): #comments coming in ui
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "MailMan Loading ..."))
        self.label_2.setText(_translate("Form", "Please give me a moment to load the modules ............"))
        self.label_3.setText(_translate("Form", ""))

class WorkerThread(QtCore.QThread):
    update_progress = QtCore.pyqtSignal(str)
    loading = QtCore.pyqtSignal(list)
    update_bar = QtCore.pyqtSignal(int)
    def __init__(self,parent=None):
        QtCore.QThread.__init__(self, parent)
    def run(self):
        self.update_progress.emit('--->>>>>>>>> Loading openpyxl <<<<<<<<<---')
        self.update_bar.emit(5)
        import openpyxl as opx
        self.loading.emit(['opx',opx])
        self.update_progress.emit('--->>>>>>>>>> Loading glob <<<<<<<<<<---')
        self.update_bar.emit(10)
        import glob
        self.loading.emit(['glob',glob])
        self.update_progress.emit('--->>>>>>>>> Loading getpass <<<<<<<<<<---')
        self.update_bar.emit(13)
        import getpass
        self.loading.emit(['getpass',getpass])
        self.update_progress.emit('--->>>>>>>>>>>> Loading sys <<<<<<<<<<<---')
        self.update_bar.emit(16)
        import sys
        self.loading.emit(['sys',sys])
        self.update_progress.emit('--->>>>>>>>> Loading pathlib <<<<<<<<<<---')
        self.update_bar.emit(19)
        from pathlib import Path
        self.loading.emit(['Path',Path])
        self.update_progress.emit('--->>>>>>>>>>>> Loading os <<<<<<<<<<<<---')
        self.update_bar.emit(24)
        import os
        self.loading.emit(['os',os])
        self.update_progress.emit('--->>>>>>>>> Loading win32ui <<<<<<<<<<---')
        self.update_bar.emit(32)
        import win32ui as wi
        self.loading.emit(['wi',wi])
        self.update_progress.emit('--->>>>> Loading win32com.client <<<<<<---')
        self.update_bar.emit(40)
        import win32com.client as wc
        self.loading.emit(['wc',wc])
        self.update_progress.emit('--->>>>>>>>> Loading Itertools <<<<<<<<---')
        self.update_bar.emit(48)
        from itertools import zip_longest
        self.loading.emit(['zip_longest',zip_longest])
        self.update_progress.emit('--->>>>>>>>>>>> Loading re <<<<<<<<<<<<---')
        self.update_bar.emit(56)
        import re
        self.loading.emit(['re',re])
        self.update_progress.emit('--->>>>>>>>>>> Loading time <<<<<<<<<<<---')
        self.update_bar.emit(64)
        import time
        self.loading.emit(['time',time])
        self.update_progress.emit('--->>>>>>>>>> Loading json <<<<<<<<<<<<---')
        self.update_bar.emit(72)
        import json
        self.loading.emit(['json',json])
        self.update_progress.emit('--->>>>>>>> Loading datetime <<<<<<<<<<---')
        self.update_bar.emit(80)
        import datetime
        self.loading.emit(['datetime',datetime])
        self.update_progress.emit('--->>>>>>>>> Loading urllib <<<<<<<<<<<---')
        self.update_bar.emit(90)
        from urllib.parse import urlparse
        self.loading.emit(['urlparse',urlparse])
        self.update_progress.emit('--->>>>>>>>>> Loading glob <<<<<<<<<<<<---')
        self.update_bar.emit(99)
        import glob
        self.loading.emit(['glob',glob])


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
