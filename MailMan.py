from PyQt5 import QtCore, QtGui, QtWidgets
from MailManConf import Ui_Conf_Window
from MailManPreprocessors import Ui_Init_Window
from Body import Ui_BodyWindow
from send_mail import SendMailThread
from archive_mail import SaveMailThread
import webbrowser

class Ui_MainWindow(object):

    def __init__(self,getpass,datetime,wc,Path,wi,os,opx,sys,json,glob,urlparse,time,re): #modules- variables present in this file
        self.getpass = getpass
        self.wc = wc
        self.datetime = datetime
        self.Path = Path
        self.wi = wi
        self.json = json
        self.os = os
        self.opx = opx
        self.sys = sys
        self.glob = glob
        self.re = re
        self.time = time
        self.busy = False
        self.urlparse = urlparse
        self.masterfile = ''
        self.mergelist = ''
        
    def modify_file_path(self,file_name): #if xlsx file is in shared drive this func converting f/w slash to b/w slash
        print("filename: ",file_name)
        if file_name.startswith("//"):
            normalised_path = self.os.path.normpath(file_name)
            file_name = normalised_path
            print("normalised path: ", self.os.path.normpath(file_name))
        print("modified file_name: ",file_name)
        return file_name
            
    def mail_store(self): #archiving a mail- open master excel sheet
        file_name = ''
        home_dir = str(self.Path.home())
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(None,'Open Master Excel file',home_dir, "xlsx files (*.xlsx)")
        print("file path from qtwidges in mail_store function:",file_name)
        if file_name:
            file_name = self.modify_file_path(file_name)
            self.masterfile = file_name
        if self.masterfile:
            self.textEdit.clear()
            self.progressBar.setValue(0) #reset progress bar and showing next string in ui
            self.busy = True
            self.worker = SaveMailThread(self.opx,self.sys,self.os,self.wi,self.wc,self.re,self.time,self.datetime,self.glob,self.Path,self.masterfile)
            self.worker.start()
            self.worker.update_progress.connect(self.update_log)
            self.worker.update_bar.connect(self.update_prog_bar)
            self.worker.finished.connect(self.evt_finished)

    def evt_finished(self):
        self.busy = False
        self.show_message('Finished','Process Finished ')

    def engine_busy(self):
        self.show_message('Busy','Can not perform the action as the engine is busy')

    def update_prog_bar(self,val):
        self.progressBar.setValue(val)

    def update_log(self,val):
        if val.startswith('Error'):
            self.error_msg('Error!!!!',val)
        self.textEdit.insertPlainText(val)

    def error_msg(self,header,body): #if any error happend while master file opening
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(header)
        msg.setIcon(QtWidgets.QMessageBox.Critical)
        msg.setInformativeText('Fatal Error Occurred')
        msg.setText(body)
        msg.exec_()

    def show_message(self,header,body):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(header)
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText(body)
        msg.exec_()

    def mail_send(self): #send file while opening master, merge file
        file_name = ''
        home_dir = str(self.Path.home())
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(None,'Open Master Excel file',home_dir, "xlsx files (*.xlsx)")
        print("file path from qtwidges in mail_send masterfile function:",file_name)
        if file_name:
            file_name = self.modify_file_path(file_name)
            self.masterfile = file_name
        #file_name_1, _ = QtWidgets.QFileDialog.getOpenFileName(None,'Open MergeList Excel file',home_dir, "xlsx files (*.xlsx)")
        file_name_1, _ = QtWidgets.QFileDialog.getOpenFileName(None,'Open MergeList Excel file')
        print("file path from qtwidges in mail_store merge file function:",file_name)
        if file_name:
            file_name = self.modify_file_path(file_name)
            self.mergelist = file_name_1
        if self.masterfile and self.mergelist:
            self.textEdit.clear()
            self.progressBar.setValue(0)
            self.busy = True
            self.worker = SendMailThread(self.opx,self.sys,self.os,self.wi,self.wc,self.re,self.time,self.json,self.datetime,self.urlparse,self.masterfile,self.mergelist)
            self.worker.start()
            self.worker.update_progress.connect(self.update_log)
            self.worker.update_bar.connect(self.update_prog_bar)
            self.worker.finished.connect(self.evt_finished)
        else:
            self.show_message('Error','Both Master excel and Merge List excel should be provided')

    def get_mailbox(self): #setup sent mail, inbox folder in outlook
        self.mail_list_box = []
        for i in self.wc.Dispatch("Outlook.Application").Session.Stores:
            self.mail_list_box.append(i.DisplayName)
        return self.mail_list_box

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(569, 480)
        MainWindow.setFixedSize(569,480)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("icons/mailman.jpeg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet("background-color: rgb(202, 233, 252);")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setGeometry(QtCore.QRect(20, 7, 531, 16))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(20, 44, 531, 91))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.lineEdit = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.lineEdit.setStyleSheet("background-color: rgb(211, 215, 207);")
        self.lineEdit.setReadOnly(True)
        self.lineEdit.setText("User Logged in - "+self.getpass.getuser()+' '+self.datetime.datetime.now().strftime("%d-%b-%Y"))
        self.lineEdit.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit.setClearButtonEnabled(False)
        self.lineEdit.setObjectName("lineEdit")
        self.verticalLayout.addWidget(self.lineEdit)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.lineEdit_3.setStyleSheet("background-color: rgb(211, 215, 207);")
        self.lineEdit_3.setReadOnly(True)
        self.lineEdit_3.setText("Contact koushik.b.dutta@accenture.com , ramya.duraisamy@accenture.com")
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.lineEdit_3.setAlignment(QtCore.Qt.AlignCenter)
        self.verticalLayout.addWidget(self.lineEdit_3)
        self.sendmail = QtWidgets.QPushButton(self.centralwidget)
        self.sendmail.setGeometry(QtCore.QRect(20, 197, 251, 31))
        self.sendmail.setStyleSheet("background-color: orange; font: bold 10pt")
        self.sendmail.setObjectName("sendmail")
        self.sendmail.clicked.connect(self.mail_send)
        self.savemail = QtWidgets.QPushButton(self.centralwidget)
        self.savemail.setGeometry(QtCore.QRect(290, 198, 261, 31))
        self.savemail.setStyleSheet("background-color: orange; font: bold 10pt")
        self.savemail.setObjectName("savemail")
        self.savemail.clicked.connect(self.mail_store)
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setGeometry(QtCore.QRect(20, 280, 541, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(20, 309, 531, 121))
        self.textEdit.setStyleSheet("background-color: rgb(211, 215, 207);")
        self.textEdit.setObjectName("textEdit")
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(20, 160, 531, 25))
        self.comboBox.setToolTipDuration(3)
        self.comboBox.setStyleSheet("background-color: rgb(211, 215, 207);")
        self.comboBox.setEditable(False)
        self.comboBox.addItems(self.get_mailbox())
        self.comboBox.setCurrentText('.Box.Data.Submissions')
        self.comboBox.setObjectName("comboBox")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(20, 260, 531, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(150, 230, 261, 20))
        self.label_2.setText("")
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(250, 140, 111, 20))
        self.label_3.setStyleSheet("font:bold 9pt")
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(510, 20, 21, 17))
        self.label_4.setText("")
        self.label_4.setPixmap(QtGui.QPixmap("icons/user-black.png"))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(50, 20, 16, 17))
        self.label_5.setText("")
        self.label_5.setPixmap(QtGui.QPixmap("icons/user-black.png"))
        self.label_5.setObjectName("label_5")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(231, 20, 141, 20))
        self.label.setStyleSheet("font:bold 11pt")
        self.label.setObjectName("label")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 569, 22))
        self.menubar.setObjectName("menubar")
        self.menuPreference = QtWidgets.QMenu(self.menubar)
        self.menuPreference.setObjectName("menuPreference")
        self.menuPreprocess = QtWidgets.QMenu(self.menubar)
        self.menuPreprocess.setObjectName("menuPreprocess")
        self.menuFont = QtWidgets.QMenu(self.menubar)
        self.menuFont.setObjectName("menuFont")
        self.menuConfiguration = QtWidgets.QMenu(self.menubar)
        self.menuConfiguration.setObjectName("menuConfiguration")
        self.menubody = QtWidgets.QMenu(self.menubar)
        self.menubody.setObjectName("menubody")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        self.menuVersion = QtWidgets.QMenu(self.menubar)
        self.menuVersion.setObjectName("menuVersion")
        self.menuexit = QtWidgets.QMenu(self.menubar)
        self.menuexit.setObjectName("menuexit")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.exit_action = QtWidgets.QAction(MainWindow)
        icon_exit = QtGui.QIcon()
        icon_exit.addPixmap(QtGui.QPixmap("icons/exit.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.exit_action.setIcon(icon_exit)
        self.exit_action.setObjectName("exit_action")
        self.exit_action.triggered.connect(self.close_application)
        self.cb_action = QtWidgets.QAction(MainWindow)
        self.ci_action = QtWidgets.QAction(MainWindow)
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("icons/changebackground.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.cb_action.setIcon(icon1)
        self.cb_action.setObjectName("cb_action")
        self.cb_action.triggered.connect(self.versionTwo)
        self.ct_action = QtWidgets.QAction(MainWindow)
        self.ci_action.setIcon(icon1)
        self.ci_action.setObjectName("ci_action")
        self.ci_action.triggered.connect(self.choose_initializers)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("icons/changetheme.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.ct_action.setIcon(icon2)
        self.ct_action.setObjectName("ct_action")
        self.ct_action.triggered.connect(self.versionTwo)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("icons/changefontfamily.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.cff_action = QtWidgets.QAction(MainWindow)
        self.cff_action.setIcon(icon3)
        self.cff_action.setObjectName("cff_action")
        self.cff_action.triggered.connect(self.versionTwo)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("icons/changefontsize.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.cfs_action = QtWidgets.QAction(MainWindow)
        self.cfs_action.setIcon(icon4)
        self.cfs_action.setObjectName("cfs_action")
        self.cfs_action.triggered.connect(self.versionTwo)
        self.Show_Help = QtWidgets.QAction(MainWindow)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("icons/question.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.Show_Help.setIcon(icon3)
        self.Show_Help.setObjectName("Show_Help")
        self.Show_Configuration = QtWidgets.QAction(MainWindow)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap("icons/gear.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.Show_Configuration.setIcon(icon4)
        self.Show_Configuration.setObjectName("Show_Configuration")
        self.Show_Configuration.triggered.connect(self.show_config)
        self.Version_Info = QtWidgets.QAction(MainWindow)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap("icons/showversion.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.Version_Info.setIcon(icon5)
        self.Version_Info.setObjectName("Version_Info")
        self.body_action = QtWidgets.QAction(MainWindow)
        icon11 = QtGui.QIcon()
        icon11.addPixmap(QtGui.QPixmap("icons/gear.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.body_action.setIcon(icon11)
        self.body_action.setObjectName("body_action")
        self.body_action.triggered.connect(self.bodyaction)
        self.menuPreference.addAction(self.cb_action)
        self.menuPreference.addAction(self.ct_action)
        self.menuPreprocess.addAction(self.ci_action)
        self.menuFont.addAction(self.cff_action)
        self.menuFont.addAction(self.cfs_action)
        self.menuConfiguration.addAction(self.Show_Configuration)
        self.menuHelp.addAction(self.Show_Help)
        self.menuVersion.addAction(self.Version_Info)
        self.menuexit.addAction(self.exit_action)
        self.menubody.addAction(self.body_action)
        self.menubar.addAction(self.menuPreference.menuAction())
        self.menubar.addAction(self.menuPreprocess.menuAction()) 
        self.menubar.addAction(self.menuFont.menuAction())
        self.menubar.addAction(self.menuConfiguration.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())
        self.menubar.addAction(self.menuVersion.menuAction())
        self.menubar.addAction(self.menuexit.menuAction())
        self.menubar.addAction(self.menubody.menuAction())
        #mainMenu.addAction(exitAct)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def versionTwo(self):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle("Version Two Warning")
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("This feature will be added in version 2.0 , stay tunned for updats")
        msg.exec_()

    def versionInfo(self):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle("Version")
        msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.setText("Version 1.0 Major , Stay Tunned for update")
        msg.exec_()

    def close_application(self):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle("Close Application?")
        msg.setText("Are you sure you want to close application?")
        msg.setIcon(QtWidgets.QMessageBox.Question)
        msg.setStandardButtons(QtWidgets.QMessageBox.Yes|QtWidgets.QMessageBox.No)
        x = msg.exec_()
        if x == QtWidgets.QMessageBox.Yes:
            sys.exit()
        else:
            pass

    def bodyaction(self):  #open browser inside this func
        webbrowser.open("bodyGenerator\\index.html",new=1)
        self.body_window = QtWidgets.QMainWindow()
        self.body_ui = Ui_BodyWindow()
        self.body_ui.setupUi(self.body_window)
        self.body_window.show()

    def show_config(self):
        self.cnf_window = QtWidgets.QMainWindow()
        self.cnf_ui = Ui_Conf_Window(self.wi,self.getpass,self.json,self.os,self.opx,self.sys)
        self.cnf_ui.setupUi(self.cnf_window)
        self.cnf_window.show()

    def choose_initializers(self):
        self.init_window = QtWidgets.QMainWindow()
        self.init_ui = Ui_Init_Window(self.wi,self.getpass,self.json,self.os,self.opx,self.sys)
        self.init_ui.setupUi(self.init_window)
        self.init_window.show()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "The Mail Man"))
        self.sendmail.setText(_translate("MainWindow", "Send Mail"))
        self.savemail.setText(_translate("MainWindow", "Archive Mail"))
        self.comboBox.setToolTip(_translate("MainWindow", "Please select the Mail Account"))
        self.label_3.setText(_translate("MainWindow", "Mail Box Using"))
        self.label.setText(_translate("MainWindow", "Hi I am Mail Man"))
        self.menuPreference.setTitle(_translate("MainWindow", "Preference"))
        self.menuFont.setTitle(_translate("MainWindow", "Font"))
        self.menubody.setTitle(_translate("MainWindow", "Setup"))
        self.menuPreprocess.setTitle(_translate("MainWindow", "Preprocess"))
        self.menuConfiguration.setTitle(_translate("MainWindow", "Configuration"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.menuVersion.setTitle(_translate("MainWindow", "Version"))
        self.menuexit.setTitle(_translate("MainWindow", "Exit"))
        self.cb_action.setText(_translate("MainWindow", "Change Background"))
        self.ct_action.setText(_translate("MainWindow", "Change Theme"))
        self.ci_action.setText(_translate("MainWindow", "Choose Initializers"))
        self.cff_action.setText(_translate("MainWindow", "Change Font Family"))
        self.cfs_action.setText(_translate("MainWindow", "Change Font Size"))
        self.Show_Help.setText(_translate("MainWindow", "Show Help"))
        self.exit_action.setText(_translate("MainWindow", "Exit App"))
        self.Show_Configuration.setText(_translate("MainWindow", "Show Configuration"))
        self.body_action.setText(_translate("MainWindow", "Body and Subject"))
        self.Version_Info.setText(_translate("MainWindow", "Version Info"))
