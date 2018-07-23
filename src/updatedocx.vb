Sub test()
    
    Dim excelApp As Excel.Application
    
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    
    Set excelApp = New Excel.Application
    Set wordApp = New Word.Application
    
    ' Load Data
    Windows("Index_data.xlsx").Activate
    Excel.ActiveWindow.Visible = Ture
    Excel.ActiveWorkbook.Sheets(2).Range("A9:B16").Select
    Selection.Copy
                
    MsgBox "Copy Data"
            
    Word.windows("Test Index.docx").Activate
    Word.ActiveWindow.Visible = True
    MsgBox Word.ActivateDocument.Name
    
    Set wordDoc = Word.ActiveDocument
    Set shp = wordDoc.Shapes(2)

    If shp.HasChart Then
        Set cht = shp.Chart
        cht.ChartData.Activate
        Set wb = cht.ChartData.WorkBook
        Set sht = wb.Sheets(1)
        
        sht.Activate
        sht.Range("A2:B2").Select
        'Insert Data'
        Selection.Insert Shift:=xlDownb
        Range("C6").Select
        
        MsgBox "Insert Data"
        cht.ChartData.Workbook.Close
        cht.Refresh
    End If

    
End Sub
