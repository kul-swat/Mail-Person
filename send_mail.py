from PyQt5 import QtCore, QtGui, QtWidgets
from urllib.parse import unquote

class SendMailThread(QtCore.QThread):
    update_progress = QtCore.pyqtSignal(str)
    update_bar = QtCore.pyqtSignal(int)

    def __init__(self,opx,sys,os,wi,wc,re,time,json,datetime,urlparse,masterfile,mergelist,isFileAttached,parent=None):
        QtCore.QThread.__init__(self, parent)
        self.opx = opx
        self.sys = sys
        self.os = os
        self.wi = wi
        self.wc = wc
        self.re = re
        self.time = time
        self.json = json
        self.datetime = datetime
        self.urlparse = urlparse
        self.masterfile = masterfile
        self.mergelist = mergelist
        self.error = False
        self.isFileAttached=isFileAttached

    def check(self): #everything is fine

        if not(self.os.path.isfile('config.prop') and self.os.stat('config.prop').st_size):
            self.error = True
            self.update_progress.emit('Error, config.prop is missing or Blank\n')
            
        if not self.os.access(r'\\uk.corporg.net\ngdfs\Shared\Cohqfs01\EVERYONE\WK24DATA\Generators',self.os.R_OK | self.os.W_OK | self.os.X_OK):
            self.error = True
            self.update_progress.emit('Error, Week24Data\Generator folder access is denied\n')

        if not(self.os.path.isfile('Body.txt') and self.os.stat('Body.txt').st_size):
            self.error = True
            self.update_progress.emit('Error, Body.txt is missing or Blank\n')

        if not(self.os.path.isfile('Subject.txt') and self.os.stat('Subject.txt').st_size):
            self.error = True
            self.update_progress.emit('Error, Subject.txt is missing or Blank\n')

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
                try:
                    self.draft_mailman = self.namespace.GetDefaultFolder(16).Folders('MailMan')
                except:
                    try:
                        self.namespace.GetDefaultFolder(16).Folders.Add('MailMan') #folders r present
                        self.draft_mailman = self.namespace.GetDefaultFolder(16).Folders('MailMan') #draft
                    except Exception as ex:
                        self.error = True
                        self.update_progress.emit('Error,Failed to create MailMan folder inside Draft inside Outlook\n')
                try:
                    self.sentbox_mailman = self.namespace.GetDefaultFolder(5).Folders('MailMan')
                except:
                    try:
                        self.namespace.GetDefaultFolder(5).Folders.Add('MailMan')
                        self.sentbox_mailman = self.namespace.GetDefaultFolder(5).Folders('MailMan')
                    except Exception as ex:
                        self.error = True
                        self.update_progress.emit('Error,Failed to create MailMan folder inside Sentbox inside Outlook\n')
            except Exception as ex:
                self.error = True
                self.update_progress.emit('Error,Failed to create outlook session dispatch\n')

        try:
            print("masterfile: ",self.masterfile)
            print("mergelist: ",self.mergelist)
            self.master_wb = self.opx.load_workbook(self.masterfile)
            self.master_ws = self.master_wb.worksheets[0]
            self.merge_wb = self.opx.load_workbook(self.mergelist)
            self.merge_ws = self.merge_wb.worksheets[0]
            if (str(self.master_ws['O'][0].value).strip()!= 'Ref No' or str(self.master_ws['A'][0].value).strip()!= 'Generator Name'
                or str(self.master_ws['U'][0].value).strip()!= 'Mail Sent On' or str(self.master_ws['V'][0].value).strip()!= 'Response received On'
                or str(self.master_ws['AI'][0].value).strip()!= 'acknowledge'):
                self.error = True
                self.update_progress.emit('Error,Master Excel Template Mismatch\n')
            if (str(self.merge_ws['L'][0].value).strip()!= 'Ref No' or str(self.merge_ws['A'][0].value).strip()!= 'Generator Name'
                or str(self.merge_ws['B'][0].value).strip()!= 'Company' or str(self.merge_ws['C'][0].value).strip()!= 'First Name'
                or str(self.merge_ws['K'][0].value).strip()!= 'Post Code' or str(self.merge_ws['E'][0].value).strip()!= 'Primary Contact Email'
                or str(self.merge_ws['F'][0].value).strip()!= 'Secondary Contact Email' or str(self.merge_ws['D'][0].value).strip()!= 'Contact'
                or str(self.merge_ws['G'][0].value).strip()!= 'Address 1' or str(self.merge_ws['H'][0].value).strip()!= 'Address 2'
                or str(self.merge_ws['I'][0].value).strip()!= 'Address 3' or str(self.merge_ws['J'][0].value).strip()!= 'Address 4'):
                self.error = True
                self.update_progress.emit('Error,Merge List Excel Template Mismatch\n')
        except Exception as ex:
            self.error = True
            self.update_progress.emit('Error,Failed to parse master file and merge list\n')

    def read_excel(self):
        common_attachment = ''
        display = ''
        review = ''
        self.update_progress.emit('Parsing config.prop\n')
        with open('config.prop','r') as config_file: #config_prop from conf window, mail-ma.py
            data = self.json.loads(config_file.read())
            common_attachment=data['common_attachment']
            display=data['display']
            review=data['review']

        rowct = 2
        maildata = []
        while type(rowct)==type(1):
            temp = dict()
            try:
                if self.merge_ws['A'+str(rowct)].value:
                    temp['Generator Name']=self.merge_ws['A'+str(rowct)].value
                    temp['Company']=self.merge_ws['B'+str(rowct)].value
                    temp['First Name']=self.merge_ws['C'+str(rowct)].value
                    temp['Contact']=self.merge_ws['D'+str(rowct)].value
                    temp['Primary Contact Email']=self.merge_ws['E'+str(rowct)].value
                    temp['Secondary Contact Email']=self.merge_ws['F'+str(rowct)].value
                    temp['Address 1']=self.merge_ws['G'+str(rowct)].value
                    temp['Address 2']=self.merge_ws['H'+str(rowct)].value
                    temp['Address 3']=self.merge_ws['I'+str(rowct)].value
                    temp['Address 4']=self.merge_ws['J'+str(rowct)].value
                    temp['Post Code']=self.merge_ws['K'+str(rowct)].value
                    temp['common_attachment']=common_attachment
                    temp['Ref No']=self.merge_ws['L'+str(rowct)].value

                    #Merge list in Shared drive, accessing it- automatic
                    if self.mergelist.startswith("//"):
                        #take the merge list path
                        print("send_mail:171  getting mergelist path:   ",self.mergelist)
                        #taking till folder where mergelist present
                        mergelist_folder=self.os.path.dirname(self.mergelist)
                        print("send_mail:174  getting mergelist folder:   ",mergelist_folder)
                        #take the file path read from hyperlink
                        #read hyperlink
                        hyperlink_path=self.urlparse(self.merge_ws['L'+str(rowct)].hyperlink.target).path
                        print("send_mail:178  hyperlink_path:   ",hyperlink_path)
                        unquote_hyperlink_path=unquote(hyperlink_path)
                        print("send_mail:180  unquote_hyperlink_path:   ",unquote_hyperlink_path)
                        #normal_hyperlink_path=self.os.path.normpath(unquote_hyperlink_path)
                        #print("send_mail:178  normal_hyperlink_path:   ",normal_hyperlink_path)
                        #join both
                        joined_hyperlink_path=self.os.path.join(mergelist_folder,unquote_hyperlink_path)
                        print("send_mail:185  joined_hyperlink_path:   ",joined_hyperlink_path)
                        #normalise the joined_hyperlink_path
                        normalised_hyperlink_path=self.os.path.normpath(joined_hyperlink_path)
                        print("send_mail:188  normalised_hyperlink_path:   ",normalised_hyperlink_path)
                        #assign the joined path in temp['File Path']
                        temp['File Path'] = normalised_hyperlink_path
                    else: #if the merge list in local drive, manual
                        p=self.urlparse(self.merge_ws['L'+str(rowct)].hyperlink.target).path
                        print("1:", self.merge_ws['L'+str(rowct)])
                        print("2:",self.merge_ws['L'+str(rowct)].hyperlink)
                        print("3:",self.merge_ws['L'+str(rowct)].hyperlink.target)
                        
                        #target_file_add = urlparse(cell_obj.hyperlink.target).path
                        print("reading hyperlink: ",p)
                        file_add=unquote(p)
                        print("unquoting: ",file_add)
                        file_add=self.os.path.normpath(file_add)[1:]
                        print("normal path: ",file_add)
                        temp['File Path'] = file_add
                        #temp_path='\\\\'+maildata['File Path']
                        
                    print("temp['File Path']:", temp['File Path'])
                    #'\\'+self.os.path.abspath(self.os.path.join(p.netloc, p.path))
                        #temp['Category']=tmpt_ws['A'+str(rowct)].value
                    temp['Review']=review
                    if display == 'YES':
                        temp['Display']=True
                    else:
                        temp['Display']=False
                    maildata.append(temp)
                    rowct = rowct + 1
                else:
                    rowct = 'EOF'
            except Exception as ex:
                self.error = True
                self.update_progress.emit('Error occurred while parsing and fetching data from Merge List\n')
        return maildata

    def mailmove(self,mail,maildata,archive):
        parent_dir = r'\\uk.corporg.net\ngdfs\Shared\Cohqfs01\EVERYONE\WK24DATA\Generators'
        filename = self.re.sub('[^A-Za-z0-9]+', ' ', mail.Subject)+' sent1.msg'
        try:
            for fldr_name in maildata['Generator Name'].split(','):
                if self.os.path.isdir(self.os.path.join(parent_dir,fldr_name)):
                    cur_year = self.datetime.datetime.now().strftime('%Y')
                    filepath = self.os.path.join(parent_dir,fldr_name,"Correspondence",cur_year)
                    if not self.os.path.isdir(filepath):
                        self.os.makedirs(filepath)
                    print("Saving Mail files")
                    mail.SaveAs(self.os.path.join(filepath,filename))
                else:
                    print("Filepath doesn't exist- creating folder")
                    filepath = self.os.path.join(parent_dir,fldr_name)
                    cur_year = self.datetime.datetime.now().strftime('%Y')
                    self.os.mkdir(filepath)
                    filepath = self.os.path.join(filepath,"Correspondence",cur_year)
                    self.os.makedirs(filepath)
                    print("Saving Mail files")
                    mail.SaveAs(self.os.path.join(filepath,filename))
                    
            mail.Move(archive)
        except Exception as ex:
            self.error = True
            self.update_progress.emit('Error, Failed to archive the mail to sentbox\n')

    def write_excel(self,maildata):
        self.update_progress.emit('Writting response back to master excel\n')
        ref_no = str(maildata['Ref No']).strip()
        gen_name = maildata['Generator Name']
        print(ref_no)
        try:           
            '''
            sendbox = self.namespace.GetDefaultFolder(5)
            item_count = sendbox.Items.Count
            for mail in sendbox.Items:
                if mail.Class == 43:
                    found = False
                    try:
                        if mail.Categories == ref_no:
                            found = True
                    except:
                        if (ref_no in mail.Subject) or (ref_no in mail.Body):
                            found = True
                    if found:
                        maildata['rcvtime']=mail.ReceivedTime.strftime('%d-%b-%Y %H:%M:%S %p')
                        maildata['convo_id']=mail.ConversationID
                        #self.mailmove(mail,maildata,self.sentbox_mailman)
                        break
            '''
            ref_no_index = [i.value for i in list(self.master_ws['O'])].index(ref_no)
            self.master_ws['U'+str(ref_no_index+1)].value = self.datetime.datetime.now().strftime('%d-%b-%Y %H:%M:%S %p')
            #self.master_ws['AI'+str(ref_no_index+1)].value = maildata['convo_id']
            self.master_wb.save(self.masterfile)
        except Exception as ex:
            self.error = True
            self.update_progress.emit('Error, error while writing back to master - '+str(ref_no)+'\n'+str(ex))

    def sendmail(self,allmaildata):
        print("Excel rows:",allmaildata)
        for maildata in allmaildata:
            print("\nMaildata:", maildata)
            self.update_progress.emit('Preparing to sendmail/savedraft for ref no '+str(maildata['Ref No']))
            mail = self.outlook.CreateItem(0)
            mail.SentOnBehalfOfName = 'data.submissions@nationalgrideso.com'
            mail.To = maildata['Primary Contact Email']
            mail.Categories = maildata['Ref No']
            try:
                f = open('Subject.txt','r')
                txt = f.read()
                print("Accessing subject file",txt)
                for cols in list(maildata.keys()):
                    rep = '##'+str(cols)+'##'
                    val = str(maildata[cols])
                    txt = txt.replace(rep,val)
                print("Replacing subject text",txt)
                mail.Subject = txt
                f.close()
                f = open('Body.txt','r')
                print("line No. 267: ", maildata)
                txt = f.read()
                for cols in list(maildata.keys()):
                    print("\n\n\n line No. 270:   ", cols)
                    rep = '##'+str(cols)+'##'
                    print("\n line No. 272:   ", rep)
                    val = str(maildata[cols])
                    txt = txt.replace(rep,val)
                    # if cols=='Ref No':
                    #     ref_no = str(maildata[cols])
                    #     txt = txt.replace('&'+str(cols)+'&',"")
                    # else:
                    #     rep = '&'+str(cols)+'&'
                    #     val = str(maildata[cols])
                    #     txt = txt.replace(rep,val)
                print("Before assigning text to html body")
                print("line No. 281:  ",txt)
                mail.Body = ""
                mail.HTMLbody = txt
                print("Replacing mail.HTMLbody text",mail.HTMLbody)
                f.close()

                #print("Body:", mail.Body)
                # print("Line: 275   Before changing to Arial:", mail.HTMLbody)
                # mail_htmlbody=mail.HTMLbody
                # mail_htmlbody=mail_htmlbody.replace("Calibri","Arial")
                # mail.HTMLbody = mail_htmlbody
                # print("Line: 279   Arial changing to Arial:", mail.HTMLbody)
                #open_sign = open('Signature.txt')
                #sign = open_sign.read()

                

                # reference_template = f'''<div style = "font-size: 14.75px;">
                #                 <FONT FACE="Arial">
                #                     {ref_no} 
                #                 </FONT>
                #             </div>'''
                # pre=mail.HTMLbody[:mail.HTMLbody.rfind("</P>")+4]
                # post=mail.HTMLbody[mail.HTMLbody.rfind("</P>")+4:]
                #mail.HTMLbody=pre+sign+reference_template+post
                # mail.HTMLbody=pre+reference_template+post
                # print("\n line:337  After inserting signature HTML Body\n", mail.HTMLbody)
                print("\n line:310  HTML Body\n", mail.HTMLbody)
                if self.mergelist.startswith("//"):
                    print("Line:312 shared drive")
                    temp_path=maildata['File Path']
                else:
                    print("Line:315 Local drive")
                    temp_path='\\\\'+maildata['File Path']

                print("File path of attachment",temp_path)
                if self.isFileAttached:
                    mail.Attachments.Add(temp_path)
                    #mail.Attachments.Add(r"\\uk.corporg.net\\ngdfs\\Shared\\Cohqfs01\\EVERYONE\\WK24DATA\\Week 24 Submissions 2020\\Correspondence & Letters 2020\\Generator PDF\\PDF 1 to 50\\1_3R Energy Solutions.pdf")
                    for fld in maildata['common_attachment'].split(','):
                        mail.Attachments.Add(fld)
                        print("folder Attachment:",fld) #generic attachment path
            except Exception as ex:
                self.error = True
                self.update_progress.emit('Error while parsing mail body and subject and attachments '+maildata['Ref No']+'\n'+str(ex))

            if not self.error:
                if maildata['Secondary Contact Email']:
                    mail.CC = maildata['Secondary Contact Email']
                mail.Display(maildata['Display'])
                if str(maildata['Review']).upper() == 'YES':
                    mail.Move(self.draft_mailman)
                else:
                    try:
                        #mail.Send()
                        maildata['resp'] = 'Success !! Mail sent successfully'
                        self.update_progress.emit('Mail Sent successfully '+str(maildata['Ref No']))
                        maildata['rcvtime']=self.datetime.datetime.now().strftime('%d-%b-%Y %H:%M:%S %p')
                    except Exception as ex:
                        maildata['resp'] = 'Error !! Error while sending mail '+str(ex)
        return allmaildata

    def run(self):
        self.update_bar.emit(2)
        self.check()
        self.update_bar.emit(10)
        if not self.error:
            maildata = self.read_excel()
            self.update_bar.emit(18)
            if not self.error:
                if maildata:
                    response=self.sendmail(maildata)
                self.update_bar.emit(62)
                if self.isFileAttached:
                    if not self.error:
                        for i in response:
                            self.write_excel(i)
                        self.update_bar.emit(100)
                    else:
                        self.update_bar.emit(100)
        else:
            self.update_bar.emit(18)
        self.update_bar.emit(100)