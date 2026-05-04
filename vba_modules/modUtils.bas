Function GetColumnIndex(ws As Worksheet, colName As String) As Long
    Dim cell As Range
    
    For Each cell In ws.Rows(1).Cells
        If Trim(LCase(cell.Value)) = Trim(LCase(colName)) Then
            GetColumnIndex = cell.Column
            Exit Function
        End If
    Next cell
    
    GetColumnIndex = -1 ' Not found
End Function
