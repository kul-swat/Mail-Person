from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Conf_Window(object):
    def __init__(self,wi,getpass,json,os,opx,sys):
        self.wi = wi
        self.getpass = getpass
        self.json = json
        self.os = os
        self.opx = opx
        self.sys = sys
    def listVal(self):
        self.val = ['YES','NO']
        return self.val
    def genAttachment(self):
        file = self.wi.CreateFileDialog(1,"All Files (*.*)|*.*|")
        file.DoModal()
        if self.lineEdit.text().strip():
            txtval = self.lineEdit.text()
            txtval = txtval + ',' + file.GetPathName()
            self.lineEdit.setText(txtval)
        else:
            self.lineEdit.setText(file.GetPathName())
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(491, 335)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 20, 101, 31))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(51, 60, 61, 31))
        self.label_2.setObjectName("label_2")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(130, 27, 261, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(130, 66, 261, 20))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_2.setText(self.getpass.getuser())
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(400, 25, 75, 23))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.genAttachment)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(78, 102, 40, 31))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(78, 132, 40, 31))
        self.label_4.setObjectName("label_4")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(30, 170, 441, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setStyleSheet("background-color: orange; font: bold 10pt")
        self.pushButton_2.clicked.connect(self.submit)
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(30, 250, 441, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.setStyleSheet("background-color: orange; font: bold 10pt")
        self.pushButton_3.clicked.connect(self.validate_prc)
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(30, 210, 441, 31))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.setStyleSheet("background-color: orange; font: bold 10pt")
        self.pushButton_4.clicked.connect(self.clear)
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(130, 110, 90, 22))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.setEditable(False)
        self.comboBox.addItems(['YES','NO'])
        self.comboBox_2 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_2.setGeometry(QtCore.QRect(130, 138, 90, 20))
        self.comboBox_2.setObjectName("comboBox_2")
        self.comboBox_2.setEditable(False)
        self.comboBox_2.addItems(['NO','YES'])
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 491, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.default_val()

    def default_val(self):
        if self.os.path.isfile('config.prop') and self.os.stat('config.prop').st_size :
            with open('config.prop') as json_file:
                data = self.json.load(json_file)
                if data:
                    self.lineEdit.setText(data['common_attachment'])
                    self.lineEdit_2.setText(data['updated_by'])
                    self.comboBox.setCurrentText(data['display'])
                    self.comboBox_2.setCurrentText(data['review'])

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "Generic Attachment"))
        self.label_2.setText(_translate("MainWindow", "Current User"))
        self.pushButton.setText(_translate("MainWindow", "Browse"))
        self.label_3.setText(_translate("MainWindow", "Display"))
        self.label_4.setText(_translate("MainWindow", "Review"))
        self.pushButton_2.setText(_translate("MainWindow", "Submit"))
        self.pushButton_3.setText(_translate("MainWindow", "Validate"))
        self.pushButton_4.setText(_translate("MainWindow", "Clear"))

    def displayMsg(self,msgs,sub):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(sub)
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText(msgs)
        msg.exec_()

    def submit(self):
        update_insert = False
        state = dict()
        state['common_attachment']=self.lineEdit.text()
        state['updated_by']=self.lineEdit_2.text()
        if self.comboBox.currentText()=='YES':
            state['display']='YES'
        else:
            state['display']='NO'
        if self.comboBox_2.currentText()=='YES':
            state['review']='YES'
        else:
            state['review']='NO'
        if self.os.path.isfile('config.prop') and self.os.stat('config.prop').st_size :
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("Overrite Config?")
            msg.setText("Are you sure you want to overwrite?")
            msg.setIcon(QtWidgets.QMessageBox.Question)
            msg.setStandardButtons(QtWidgets.QMessageBox.Yes|QtWidgets.QMessageBox.No)
            x = msg.exec_()
            if x == QtWidgets.QMessageBox.Yes:
                update_insert = True

        elif self.os.path.isfile('config.prop') and (not self.os.stat('config.prop').st_size) :
            update_insert = True
        else:
            update_insert = True

        if update_insert:
            try:
                with open('config.prop', 'w+') as outfile:
                    self.json.dump(state, outfile)
                self.displayMsg("Config file updated successfully",'SUCCESS!!')
                self.lineEdit.setText("")
                self.lineEdit_2.setText("")
            except Exception as ex:
                self.displayMsg(str(ex),'ERROR')

    def clear(self):
        self.lineEdit.setText("")
        self.lineEdit_2.setText("")

    def validate_prc(self):
        self.displayMsg('Validate Merge list has been initiated , It Might take some time 2-10 minutes for openpyxl engine','Information')
        self.statusbar.showMessage('Using openpyxl Engine')
        self.statusbar.showMessage('Importing Done')
        mrgList = self.wi.CreateFileDialog( 1, ".xlsx", "", 0, "xlsx Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*|")
        mrgList.DoModal()
        valFile = ''
        if not mrgList.GetPathName().strip():
            self.statusbar.showMessage('No xlsx file chosen !! Sorry')
        else:
            self.statusbar.showMessage('Validation initiated')
            valFile = mrgList.GetPathName()
            workbook = self.opx.load_workbook(valFile)
            worksheet = workbook.worksheets[0]
            rowct = 2
            data = []
            filePath = r'\\uk.corporg.net\ngdfs\Shared\Cohqfs01\EVERYONE\WK24DATA\Generators'
            while type(rowct)==type(1):
                temp = dict()
                try:
                    if worksheet['A'+str(rowct)].value:
                        status = ''
                        fldr_list = worksheet['A'+str(rowct)].value.split(',')
                        for fli in fldr_list:
                            if self.os.path.isdir(self.os.path.join(filePath,fli)):
                                status = status+'Success '+fli+'||'
                            else:
                                status = status+'Error '+fli+'||'
                        worksheet['O'+str(rowct)].value = status
                        rowct = rowct + 1
                    else:
                        rowct = 'EOF'
                except Exception as ex:
                    return data
                    print('Error occurred '+str(ex))
                    self.sys.exit(1)
            workbook.save(valFile)
            self.statusbar.showMessage('Validation process finished')
