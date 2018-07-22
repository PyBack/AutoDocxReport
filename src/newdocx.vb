Public Sub NewWordApp()

    'Create variables to reference objects
    '(This line is not needed in Python; you don't need to declare variables 
    'or their types before using them)
    Dim wordApp As Word.Application, wordDoc As Word.Document

    'Create a new instance of a Word Application object
    '(Another difference - in VBA you use Set for objects and simple assignment for 
    'primitive values. In Python, you use simple assignment for objects as well.)
    Set wordApp = New Word.Application

    'Show the application
    wordApp.Visible = True

    'Create a new document in the application
    Set wordDoc = wordApp.Documents.Add()

    'Set the text of the first paragraph
    '(A Paragraph object doesn't have a Text property. Instead, it has a Range property
    'which refers to a Range object, which does have a Text property.)
    wordDoc.Paragraphs(1).Range.Text = "Hello, World!"

End Sub
