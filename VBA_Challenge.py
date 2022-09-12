Sub Test()
Dim WorksheetsCount As Integer
Dim I As Integer
WorksheetsCount = ActiveWorkbook.Worksheets.Count

For I = 1 To WorksheetsCount
    'Get current worksheet
    Dim currentWorksheet As Worksheet
    Set currentWorksheet = ActiveWorkbook.Worksheets(I)
    
    'Clear previous result
    currentWorksheet.Columns("I:L").ClearContents
    
    'Generate Ticker data
    Dim rowNum As Long
    rowNum = currentWorksheet.Cells(currentWorksheet.Rows.Count, "A").End(xlUp).Row
    currentWorksheet.Range("A1:A" & rowNum).AdvancedFilter _
    Action:=xlFilterCopy, CopyToRange:=currentWorksheet.Range("I1"), _
    Unique:=True
    
    'Set Ticker column header
    currentWorksheet.Range("I1").Value = "Ticker"
    
    'Generate Yearly Change data
    currentWorksheet.Range("J1").Value = "Yearly Change"
    
    Dim O As Integer
    Dim Top As Long
    Dim Bottom As Long
    Dim tickerSymbol As String
    Dim yearlyChange As Double
    For O = 2 To currentWorksheet.Cells(currentWorksheet.Rows.Count, "I").End(xlUp).Row
        tickerSymbol = currentWorksheet.Range("I" & O).Value
        Top = currentWorksheet.Range("A:G").Find(What:=tickerSymbol).Row
        Bottom = currentWorksheet.Range("A:G").Find(What:=tickerSymbol, _
        SearchDirection:=xlPrevious).Row
        
        yearlyChange = currentWorksheet.Range("F" & Bottom).Value _
        - currentWorksheet.Range("C" & Top).Value
        currentWorksheet.Range("J" & O).Value = yearlyChange
        
        'Set cell color, red if negtive, green if positive
        If yearlyChange >= 0 Then
            currentWorksheet.Range("J" & O).Interior.ColorIndex = 4
        Else
            currentWorksheet.Range("J" & O).Interior.ColorIndex = 3
        End If
        
        'Generate Percent Change
        currentWorksheet.Range("K" & O).Value = yearlyChange / currentWorksheet.Range("C" & Top).Value
        'Format Percentage
        currentWorksheet.Range("K" & O).NumberFormat = "0.00%"
        
        'Generate Total Stock Volume
        currentWorksheet.Range("L" & O).Value = Application.WorksheetFunction.SumIf(currentWorksheet.Range("A2:A" & rowNum), tickerSymbol, _
        currentWorksheet.Range("G2:G" & rowNum))
        Next O
    currentWorksheet.Range("K1").Value = "Percent Change"
    currentWorksheet.Range("L1").Value = "Total Stock Volume"
    currentWorksheet.Columns("I:L").AutoFit
Next I
End Sub

