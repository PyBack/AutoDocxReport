Sub test()

    Dim excelApp As Excel.Application
    Dim wb As Excel.Workbook
    Dim sht As Excel.Worksheet
    
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    
    Dim shp As InlineShape
    Dim cht As Word.Chart
    
    Set wordApp = New Word.Application
    wordApp.Visible = True
    
    Set wordDoc = Word.Documents.Open("C:\Users\assa\Downloads\테스트.docx")
    
    
    For Each shp In ActiveDocument.InlineShapes
        If shp.HasChart Then
            Set cht = shp.Chart
            'Here comes the Question: how to assign the chartdata.workbook to wb?
        
            ' Set ils = ActiveDocument.InlineShapes(Index)
            ' Set c = ils.Chart
            cht.ChartData.Activate
        
            Set wb = cht.ChartData.Workbook
            Set sht = wb.Worksheets(1)
            ' Set lo = ws.ListObjects(1)
            ' lo.Resize wb.Application.Range("A1:D7")
            ' ws.Cells(6, 1).Value = "New category"
            ' ws.Cells(6, 2).Value = 6.8

            cht.ChartData.Workbook.Sheets(1).Cells(1, 1).Value = "date"
            cht.ChartData.Workbook.Sheets(1).Cells(1, 2).Value = 10

            cht.ChartData.Workbook.Close

            cht.Refresh

            Set cht = Nothing
            ' Set shp = Nothing
        End If
    Next shp
    
    Set ils = Nothing

End Sub
