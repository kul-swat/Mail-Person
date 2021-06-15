#!/usr/bin/env python
from PyQt5 import QtCore, QtGui, QtWidgets

class SaveMailThread(QtCore.QThread):

    update_progress = QtCore.pyqtSignal(str)
    update_bar = QtCore.pyqtSignal(int)

    def __init__(self,opx,sys,os,wi,wc,re,time,datetime,glob,Path,masterfile,parent=None):
        QtCore.QThread.__init__(self, parent) #archive the mail
        self.opx = opx
        self.sys = sys
        self.os = os
        self.wi = wi
        self.wc = wc
        self.re = re
        self.time = time
        self.glob = glob
        self.datetime = datetime
        self.Path = Path
        self.masterfile = masterfile
        self.error = False

    def check(self):#check outlook

        if not self.os.access(r'\\uk.corporg.net\ngdfs\Shared\Cohqfs01\EVERYONE\WK24DATA\Generators',self.os.R_OK | self.os.W_OK | self.os.X_OK):
            self.error = True
            self.update_progress.emit('Error, Week24Data\Generator folder access is denied\n')

        try:
            self.wi.FindWindow(None, "Microsoft Outlook")
        except self.wi.error:
            try:
                self.os.startfile("outlook")
            except Exception as ex:
                self.error = True
                self.update_progress.emit('Error,Failed to start outlook\n')

        if not self.error:
            try:
                self.outlook = self.wc.Dispatch('outlook.application')
                self.namespace = self.outlook.GetNamespace('MAPI')
            except Exception as ex:
                self.error = True
                self.update_progress.emit('Error,Failed to initiate MAPI outlook session\n')
            if not self.error:
                try:
                    self.sentbox_mailman = self.namespace.GetDefaultFolder(5).Folders('MailMan')
                except:
                    try:
                        self.namespace.GetDefaultFolder(5).Folders.Add('MailMan')
                        self.sentbox_mailman = self.namespace.GetDefaultFolder(5).Folders('MailMan') #sent box
                    except Exception as ex:
                        self.error = True
                        self.update_progress.emit('Error,Failed to create MailMan folder inside Sentbox inside Outlook\n')
                try:
                    self.inbox_mailman = self.namespace.GetDefaultFolder(6).Folders('MailMan')#inbox
                except:
                    try:
                        self.namespace.GetDefaultFolder(6).Folders.Add('MailMan')
                        self.inbox_mailman = self.namespace.GetDefaultFolder(6).Folders('MailMan')
                    except Exception as ex:
                        self.error = True
                        self.update_progress.emit('Error,Failed to create MailMan folder inside Inbox inside Outlook\n')
                try:
                    folder=self.namespace.Folders.Item('.Box.Data.Submissions')
                    self.inbox=folder.Folders.Item("Inbox")
                except Exception as ex:
                    self.error = True
                    self.update_progress.emit('Error,Failed to access Inbox of .Box.Data.Submissions inside Outlook\n')

                try:
                    self.sentbox=self.namespace.GetDefaultFolder(5)
                except Exception as ex:
                    self.error = True
                    self.update_progress.emit('Error,Failed to access Sent of User inside Outlook\n')

        try:
            self.master_wb = self.opx.load_workbook(self.masterfile)#
            self.master_ws = self.master_wb.worksheets[0]
            if (str(self.master_ws['O'][0].value).strip()!= 'Ref No' or str(self.master_ws['A'][0].value).strip()!= 'Generator Name'
                or str(self.master_ws['U'][0].value).strip()!= 'Mail Sent On' or str(self.master_ws['V'][0].value).strip()!= 'Response received On'
                or str(self.master_ws['B'][0].value).strip()!= 'Company' or str(self.master_ws['F'][0].value).strip()!= 'Primary Contact Email'
                or str(self.master_ws['D'][0].value).strip()!= 'First Name'):
                self.error = True
                self.update_progress.emit('Error,Master Excel Template Mismatch\n')
        except Exception as ex:
            self.error = True
            self.update_progress.emit('Error,Failed to parse master file and merge list\n')

    def read_excel(self):
        #print("Reading Mastersheet Excel workbook")
        self.update_progress.emit('parsing '+str(self.masterfile)+' to fetch value')
        rowct = 2
        data = []
        while type(rowct)==type(1):
            temp = dict()
            try:
                if self.master_ws['A'+str(rowct)].value:
                    temp['Generator Name']=self.master_ws['A'+str(rowct)].value
                    temp['Company']=self.master_ws['B'+str(rowct)].value
                    temp['First Name']=self.master_ws['D'+str(rowct)].value
                    temp['Contact']=self.master_ws['E'+str(rowct)].value
                    temp['Primary Contact Email']=self.master_ws['F'+str(rowct)].value
                    temp['Secondary Contact Email']=self.master_ws['G'+str(rowct)].value
                    temp['Address 1']=self.master_ws['H'+str(rowct)].value
                    temp['Address 2']=self.master_ws['I'+str(rowct)].value
                    temp['Address 3']=self.master_ws['J'+str(rowct)].value
                    temp['Address 4']=self.master_ws['K'+str(rowct)].value
                    temp['Post Code']=self.master_ws['L'+str(rowct)].value
                    temp['Location']=self.master_ws['M'+str(rowct)].value
                    temp['Ref No']=self.master_ws['O'+str(rowct)].value
                    temp['Mail Sent']=self.master_ws['U'+str(rowct)].value
                    temp['Convo ID']=self.master_ws['AI'+str(rowct)].value
                    data.append(temp)
                    rowct = rowct + 1
                else:
                    rowct = 'EOF'
            except Exception as ex:
                self.error = True
                self.update_progress.emit('Error occurred while parsing '+str(self.masterfile)+'\n'+str(ex))
        return data

    def mailmove(self,mail,maildata,archive,chs):
        self.update_progress.emit('Archiving process start for '+maildata['Ref No'])
        parent_dir = r'\\uk.corporg.net\ngdfs\Shared\Cohqfs01\EVERYONE\WK24DATA\Generators'
        for fldr_name in maildata['Generator Name'].split(','):
            year = self.datetime.datetime.now().year
            filepath = ''
            if self.os.path.isdir(self.os.path.join(parent_dir,fldr_name,str(year),"Correspondence")):
                filepath = self.os.path.join(parent_dir,fldr_name,str(year),"Correspondence")
                print(filepath)
            else:
                filepath = self.os.path.join(parent_dir,fldr_name,str(year),"Correspondence")
                print(filepath)
                try:
                    print("\nSaving Archive mails in folder")
                    self.Path(filepath).mkdir(parents=True,exist_ok=True)
                except Exception as ex:
                    print(str(ex))
                    self.error = True
                    self.update_progress.emit('Error,No directory found by the name of Generator or could not create it '+str(maildata['Ref No'])+'\n'+str(ex))
            if not self.error:
                files = self.glob.glob(filepath+"/*.msg")
                if files:
                    idtag = []
                    for i in files:
                        query = chs+'(\d+).msg'
                        res = self.re.findall(query,i)
                        if res:
                            idtag.append(int(res[0]))
                    if idtag:
                        filename = self.re.sub('[^A-Za-z0-9]+', ' ', mail.Subject)+' '+chs+str(max(idtag)+1)+'.msg'
                        mail.SaveAs(self.os.path.join(filepath,filename))
                    else:
                        filename = self.re.sub('[^A-Za-z0-9]+', ' ', mail.Subject)+' '+chs+'1.msg'
                        mail.SaveAs(self.os.path.join(filepath,filename))
                else:
                    filename = self.re.sub('[^A-Za-z0-9]+', ' ', mail.Subject)+' '+chs+'1.msg'
                    mail.SaveAs(self.os.path.join(filepath,filename))
        if not self.error:
            mail.Move(archive)

    def write_resp_date(self,convo_id,rcv_time,chs,ref_no):
        column = self.master_ws['O']  # Column
        column_list = [column[x].value for x in range(len(column))]
        if ref_no in column_list:
            row = column_list.index(ref_no)
            if chs=='rcv':
                self.master_ws['V'+str(row+1)].value=str(self.master_ws['V'+str(row+1)].value)+','+rcv_time
            else:
                self.master_ws['U'+str(row+1)].value=str(self.master_ws['U'+str(row+1)].value)+','+rcv_time
            if convo_id:
                self.master_ws['AI'+str(row+1)].value=convo_id

        self.master_wb.save(self.masterfile)

    def run(self):
        self.update_progress.emit('Starting Archiving mailbox')
        self.update_bar.emit(2)
        self.check()
        if not self.error:
            self.update_bar.emit(12)
            maildata = self.read_excel()
            self.update_bar.emit(20)
            if not self.error:
                upd_bar = len(maildata)/70
                progress = 0
                if maildata:
                    for data in maildata:
                        progress = progress+upd_bar
                        self.update_bar.emit(progress)
                        # Change it later
                        ref_no = data['Ref No']
                        convo_id = str(data['Convo ID']).strip()
                        if ref_no:
                            ct = self.inbox.Items.Count
                            for i in range(ct,0,-1):
                                mail = self.inbox.Items[i-1]
                                if mail.Class == 43:
                                    if (ref_no in mail.Subject or ref_no in mail.Body or ref_no in mail.Categories):
                                        print('Found in inbox')
                                        rcv_time = mail.ReceivedTime.strftime('%d-%b-%Y %H:%M:%S %p')
                                        self.mailmove(mail,data,self.inbox_mailman,'rcv')
                                        self.write_resp_date(mail.ConversationId,rcv_time,'rcv',ref_no)
                                        break
                            ct = self.sentbox.Items.Count
                            for i in range(ct,0,-1):
                                mail = self.sentbox.Items[i-1]
                                if mail.Class == 43:
                                    if (ref_no in mail.Subject or ref_no in mail.Body or ref_no in mail.Categories):
                                        print('Found in sent box')
                                        rcv_time = mail.ReceivedTime.strftime('%d-%b-%Y %H:%M:%S %p')
                                        self.mailmove(mail,data,self.sentbox_mailman,'sent')
                                        self.write_resp_date(mail.ConversationId,rcv_time,'sent',ref_no)
                                        break
                else:
                    self.update_bar.emit(100)
            else:
                self.update_bar.emit(100)
        else:
            self.update_bar.emit(100)
        self.update_bar.emit(100)
