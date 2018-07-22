import win32com.client

#Create an instance of Word.Application
wordApp = win32com.client.Dispatch('Word.Application')

#Show the application
wordApp.Visible = True

#Create a new document in the application
wordDoc = wordApp.Documents.Add()

#Set the text of the first paragraph
wordDoc.Paragraphs(1).Range.Text = "Hello, World!"
