Sub CleanData()

    Dim wsRaw As Worksheet
    Dim wsClean As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set wsRaw = ThisWorkbook.Sheets("Raw_Data")
    Set wsClean = ThisWorkbook.Sheets("Clean_Data")
    
    ' Clear old data
    wsClean.Cells.Clear
    
    ' Copy data
    wsRaw.UsedRange.Copy wsClean.Range("A1")
    
    lastRow = wsClean.Cells(wsClean.Rows.Count, 1).End(xlUp).Row
    lastCol = wsClean.Cells(1, wsClean.Columns.Count).End(xlToLeft).Column
    
    ' Remove extra spaces
    Dim i As Long, j As Long
    For i = 2 To lastRow
        For j = 1 To lastCol
            If VarType(wsClean.Cells(i, j).Value) = vbString Then
                wsClean.Cells(i, j).Value = Trim(wsClean.Cells(i, j).Value)
            End If
        Next j
    Next i
    
    MsgBox "Data Cleaned Successfully!", vbInformation

End Sub

Sub ValidateColumns()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim requiredCols As Variant
    requiredCols = Array("EmployeeID", "Department", "Attrition", "Salary")
    
    Dim i As Integer
    For i = LBound(requiredCols) To UBound(requiredCols)
        If GetColumnIndex(ws, CStr(requiredCols(i))) = -1 Then
            MsgBox "Missing Column: " & requiredCols(i), vbCritical
            Exit Sub
        End If
    Next i
    
    MsgBox "All required columns present!", vbInformation

End Sub
