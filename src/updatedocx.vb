Sub test()
    
    Dim excelApp As Excel.Application
    Dim wb As Excel.Workbook
    Dim sht As Excel.Worksheet
    
    
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    
    Dim cht As Word.Chart
    
    Set excelApp = New Excel.Application
    Set wordApp = New Word.Application
       
    ' Load Data
    'excelApp.Visible = True
    Excel.Windows("index_data.xlsx").Activate
    'Excel.ActiveWindow.Visible = Ture
    'Excel.ActiveWorkbook.Sheets(2).Range("A2:B9").Select
    'excelApp.Visible = True
    Set wb = Excel.ActiveWorkbook
    Set sht = wb.Worksheets(2)
    sht.Activate
    sht.Range("A2:B9").Select
    'excelApp.ActiveWorkbook.Sheets(2).Range("A2:B9").Select
    Selection.Copy
                
    MsgBox "Copy Data"
    
    Word.Windows("Test Index.docx").Activate
    'Word.ActiveWindow.Visible = True
    
    Set wordDoc = Word.ActiveDocument
    MsgBox wordDoc.Name
    
    'Set shp = wordDoc.Shapes(1)
    Set shp = wordDoc.InlineShapes.Item(0)
        
    ' Set shp = wordDoc.Shapes(1)

    If shp.HasChart Then
    
        Set cht = shp.Chart
        cht.ChartData.Activate
        Set wb = cht.ChartData.Workbook
        Set sht = wb.Sheets(1)
        
        sht.Activate
        sht.Range("A2:B2").Select
        
        'Insert Data'
        Selection.Insert Shift:=xlDown
        Range("C6").Select
        
        MsgBox "Insert Data"
        cht.ChartData.Workbook.Close
        cht.SetSourceData Source:="='Sheet1'!$A$1:$B$9"
        cht.Refresh
        
    End If
    
End Sub

Sub update_chart()

    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.SetSourceData Source:=Range("A1:B25")
    
End Sub
