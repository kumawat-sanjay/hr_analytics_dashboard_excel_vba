Sub ResetDashboard()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Clear only data areas
    ws.Range("B3:B5").ClearContents
    ws.Range("B7").ClearContents
    ws.Range("B10:B11").ClearContents
    ws.Range("D3:E100").ClearContents
    
    ' Clear search output
    ws.Range("H2:H8").ClearContents
    
    ' Delete charts only
    Dim ch As ChartObject
    For Each ch In ws.ChartObjects
        ch.Delete
    Next ch
    
    MsgBox "Dashboard reset (UI preserved) ??", vbInformation

End Sub
