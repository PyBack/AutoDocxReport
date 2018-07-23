Sub test()
    
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    
    Set wordApp = New Word.Application
    Word.windows("Test Index.docx").Activate

    'MsgBox Word.ActivateDocument.Name
    Set wordDoc = Word.ActiveDocument
    Set shp = wordDoc.Shapes(2)

    If shp.HasChart Then
        Set cht = shp.Chart
        cht.ChartData.Activate
        Set wb = cht.ChartData.WorkBook
        Set sht = wb.Sheets(1)
        
        MsgBox 'Open Chart Data'
        
        cht.ChartData.Close
    End If

    
End Sub
