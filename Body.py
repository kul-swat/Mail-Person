# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt5 UI code generator 5.12.3
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_BodyWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(507, 415)
        MainWindow.setStyleSheet("background-color: rgb(220, 220, 220);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(140, 10, 250, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setStyleSheet("color: rgb(85, 170, 127);")
        self.label.setObjectName("label")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(20, 40, 460, 10))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(20, 64, 120, 13))
        font = QtGui.QFont()
        font.setFamily("MS UI Gothic")
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setGeometry(QtCore.QRect(20, 90, 460, 50))
        self.plainTextEdit.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(20, 160, 100, 20))
        font = QtGui.QFont()
        font.setFamily("MS UI Gothic")
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.plainTextEdit_2 = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit_2.setGeometry(QtCore.QRect(20, 190, 460, 130))
        self.plainTextEdit_2.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.plainTextEdit_2.setObjectName("plainTextEdit_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(250, 340, 230, 30))
        self.pushButton.setStyleSheet("background-color: rgb(87, 175, 130);")
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.save_all)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 507, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Body and Subject"))
        self.label.setText(_translate("MainWindow", "MailMan Body and Subject Generator"))
        self.label_2.setText(_translate("MainWindow", "Subject of The Mail"))
        self.label_3.setText(_translate("MainWindow", "Body of The Mail"))
        self.pushButton.setText(_translate("MainWindow", "Submit"))

    def save_all(self):
        subject = self.plainTextEdit.toPlainText()
        body = self.plainTextEdit_2.toPlainText()
        buttonreply = 'GO FOR IT'
        #if not(subject.strip() and body.strip()):
            #buttonreply = QtWidgets.QMessageBox.question(self,'Save ??','Do you want to proceed?',QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,QtWidgets.QMessageBox.Yes)
        if 1==1:#buttonreply == QtWidgets.QMessageBox.Yes or buttonreply == 'GO FOR IT':
            with open('Body.txt','w+') as f:
                f.write(self.plainTextEdit_2.toPlainText())
            with open('Subject.txt','w+') as f:
                f.write(self.plainTextEdit.toPlainText())
            self.plainTextEdit.clear()
            self.plainTextEdit_2.clear()
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("DONE!!")
            msg.setIcon(QtWidgets.QMessageBox.Information)
            msg.setText("Process Finished")
            msg.exec_()
