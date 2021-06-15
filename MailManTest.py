# -*- coding: utf-8 -*-
"""
Created on Fri May 14 17:45:33 2021

@author: swati.bohidar
"""

import unittest  
from MailMan import Ui_MainWindow #(mail_store, mail_send, get_mailbox, modify_file_path, close_application, bodyaction, show_config )
from archive_mail import SaveMailThread #(mailmove, write_resp_date, run)
from MailManConf import Ui_Conf_Window #(genAttachment, listVal, default_val, displayMsg, submit, clear, validate_prc)
from send_mail import SendMailThread #(read_excel, mailmove, write_excel, sendmail, run)

import os
import openpyxl as opx


class MailmanTests(unittest.TestCase):
    
    def setUp(self):
        self.obj_uimainwindow = Ui_MainWindow(*[None]*5,os,*[None]*7)
        self.obj_SaveMailThread = SaveMailThread(*[None]*12)
        self.obj_SaveMailThread.master_wb = opx.load_workbook('MasterDemo.xlsx')#
        self.obj_SaveMailThread.master_ws = self.obj_SaveMailThread.master_wb.worksheets[0]
        self.data_output = self.obj_SaveMailThread.read_excel()
    
    def test_func_sharedDrive(self):
        output = self.obj_uimainwindow.modify_file_path("//uk.corporg.net/ngdfs/Shared/Cohqfs01/EVERYONE/WK24DATA")
        self.assertEqual(output,"\\\\uk.corporg.net\\ngdfs\\Shared\\Cohqfs01\\EVERYONE\\WK24DATA")
        
    def test_func_localDrive(self):
        output = self.obj_uimainwindow.modify_file_path("H:/MailMan")
        self.assertEqual(output,"H:/MailMan")
    
    #def test_messg(self):
        #self.obj_uimainwindow.evt_finished()
        #self.assertEqual(self.obj_uimainwindow.busy, False)
    
    def test_read_excel1(self):
        print("Generator Name: ",self.data_output[0]['Generator Name'])
        self.assertEqual(self.data_output[0]['Generator Name'],'WK24 GEN1')
    def test_read_excel2(self):
        print("Company: ",self.data_output[0]['Company'])
        self.assertEqual(self.data_output[0]['Company'],'NG ESO')
    def test_read_excel3(self):
        print("First Name: ",self.data_output[0]['First Name'])
        self.assertEqual(self.data_output[0]['First Name'],'Lilian')
    def test_read_excel4(self):
        print("Primary Contact Email: ",self.data_output[0]['Primary Contact Email'])
        self.assertEqual(self.data_output[0]['Primary Contact Email'],'paul.mathew2@nationalgrideso.com')
    def test_read_excel5(self):
        print("Location: ",self.data_output[0]['Location'])
        self.assertEqual(self.data_output[0]['Location'],'England and Wales')
    def test_read_excel6(self):
        print("Ref No: ",self.data_output[0]['Ref No'])
        self.assertEqual(self.data_output[0]['Ref No'],'WK24-GEN-20-010')
    def test_read_excel7(self):
        print("Contact: ",self.data_output[0]['Contact'])
        self.assertEqual(self.data_output[0]['Contact'],'Ms S Tonge')
    def test_read_excel8(self):
        print("Secondary Contact Email: ", self.data_output[0]['Secondary Contact Email'])
        self.assertEqual(self.data_output[0]['Secondary Contact Email'],None)
        

if __name__=="__main__":
    unittest.main()

