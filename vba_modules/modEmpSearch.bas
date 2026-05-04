Sub SearchEmployee()

    Dim wsData As Worksheet, wsDash As Worksheet
    Set wsData = ThisWorkbook.Sheets("Clean_Data")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    ' Get input
    Dim empID As String
    empID = Trim(wsDash.Range("H2").Value)
    
    If empID = "" Then
        MsgBox "Please enter Employee ID", vbExclamation
        Exit Sub
    End If
    
    ' Get column indexes
    Dim colID As Long, colName As Long, colDept As Long
    Dim colSalary As Long, colAttr As Long, colRisk As Long
    
    colID = GetColumnIndex(wsData, "EmployeeID")
    colName = GetColumnIndex(wsData, "EmployeeName")
    colDept = GetColumnIndex(wsData, "Department")
    colSalary = GetColumnIndex(wsData, "Salary")
    colAttr = GetColumnIndex(wsData, "Attrition")
    colRisk = GetColumnIndex(wsData, "RiskScore")
    
    ' Validate critical column
    If colID = -1 Then
        MsgBox "EmployeeID column not found!", vbCritical
        Exit Sub
    End If
    
    ' Find last row
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, colID).End(xlUp).Row
    
    ' Define search range (excluding header)
    Dim searchRange As Range
    Set searchRange = wsData.Range(wsData.Cells(2, colID), wsData.Cells(lastRow, colID))
    
    ' Perform MATCH (fast lookup)
    Dim matchIndex As Variant
    matchIndex = Application.Match(empID, searchRange, 0)
    
    ' If not found
    If IsError(matchIndex) Then
        MsgBox "Employee not found", vbExclamation
        wsDash.Range("H4:H8").ClearContents
        Exit Sub
    End If
    
    ' Convert to actual row number
    Dim resultRow As Long
    resultRow = matchIndex + 1   ' because search starts at row 2
    
    ' Output results safely
    If colName <> -1 Then wsDash.Range("H4").Value = wsData.Cells(resultRow, colName).Value
    If colDept <> -1 Then wsDash.Range("H5").Value = wsData.Cells(resultRow, colDept).Value
    If colSalary <> -1 Then wsDash.Range("H6").Value = wsData.Cells(resultRow, colSalary).Value
    If colAttr <> -1 Then wsDash.Range("H7").Value = wsData.Cells(resultRow, colAttr).Value
    
    If colRisk <> -1 Then
        wsDash.Range("H8").Value = wsData.Cells(resultRow, colRisk).Value
    Else
        wsDash.Range("H8").Value = "N/A"
    End If

End Sub
