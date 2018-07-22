# -*- coding: utf-8 -*-

import win32com.client

#Create an instance of Word.Application
wordApp = win32com.client.Dispatch('Word.Application')

#Show the application
wordApp.Visible = True

#Open document in the application
file_path = r"C:\Users\assa\Downloads"
file_name = u"마켓 리뷰.docx"
docx_path = file_path + "\\" + file_name
wordDoc = wordApp.Documents.Open(docx_path)


