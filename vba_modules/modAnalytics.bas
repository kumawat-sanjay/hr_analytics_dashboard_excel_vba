Sub CalculateAttrition()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim colAttrition As Long
    colAttrition = GetColumnIndex(ws, "Attrition")
    If colAttrition = -1 Then Exit Sub
    
    Dim i As Long, attritionCount As Long
    
    For i = 2 To lastRow
        If ws.Cells(i, colAttrition).Value = "Yes" Then
            attritionCount = attritionCount + 1
        End If
    Next i
    
    Dim totalEmployees As Long
    totalEmployees = lastRow - 1
    If totalEmployees = 0 Then Exit Sub
    
    Dim rate As Double
    rate = attritionCount / totalEmployees
    
    With ThisWorkbook.Sheets("Dashboard")
        .Range("B3").Value = totalEmployees
        .Range("B4").Value = attritionCount
        .Range("B5").Value = rate
        .Range("B5").NumberFormat = "0.00%"
    End With

End Sub


Sub DepartmentRisk()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim colDept As Long, colAttr As Long
    colDept = GetColumnIndex(ws, "Department")
    colAttr = GetColumnIndex(ws, "Attrition")
    
    If colDept = -1 Or colAttr = -1 Then Exit Sub
    
    Dim dictTotal As Object, dictAttr As Object
    Set dictTotal = CreateObject("Scripting.Dictionary")
    Set dictAttr = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long, dept As String
    
    For i = 2 To lastRow
        
        dept = ws.Cells(i, colDept).Value
        
        dictTotal(dept) = dictTotal(dept) + 1
        
        If ws.Cells(i, colAttr).Value = "Yes" Then
            dictAttr(dept) = dictAttr(dept) + 1
        End If
        
    Next i
    
    Dim rowOut As Long
    rowOut = 3
    
    With ThisWorkbook.Sheets("Dashboard")
        
        .Range("D3:E100").ClearContents
        .Range("E:E").NumberFormat = "0.00%"
        
        Dim key As Variant, rate As Double
        
        For Each key In dictTotal.Keys
            
            If dictTotal(key) > 0 Then
                rate = dictAttr(key) / dictTotal(key)
            Else
                rate = 0
            End If
            
            .Cells(rowOut, 4).Value = key
            .Cells(rowOut, 5).Value = rate
            
            rowOut = rowOut + 1
        Next key
        
    End With

End Sub



Sub GenerateInsights()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    Dim rate As Double
    rate = ws.Range("B5").Value
    
    If rate > 0.2 Then
        ws.Range("B7").Value = "High Attrition Risk - Immediate Action Required"
    ElseIf rate > 0.1 Then
        ws.Range("B7").Value = "Moderate Risk - Monitor Closely"
    Else
        ws.Range("B7").Value = "Workforce Stable"
    End If

End Sub



Sub SalaryAttritionAnalysis()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim colSalary As Long, colAttr As Long
    colSalary = GetColumnIndex(ws, "Salary")
    colAttr = GetColumnIndex(ws, "Attrition")
    
    If colSalary = -1 Or colAttr = -1 Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim lowAttr As Long, totalLow As Long
    Dim i As Long
    
    For i = 2 To lastRow
        
        If ws.Cells(i, colSalary).Value < 50000 Then
            
            totalLow = totalLow + 1
            
            If ws.Cells(i, colAttr).Value = "Yes" Then
                lowAttr = lowAttr + 1
            End If
            
        End If
        
    Next i
    
    Dim rate As Double
    If totalLow > 0 Then rate = lowAttr / totalLow
    
    With ThisWorkbook.Sheets("Dashboard")
        .Range("B10").Value = rate
        .Range("B10").NumberFormat = "0.00%"
    End With

End Sub



Sub ExperienceRisk()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim colExp As Long, colAttr As Long
    colExp = GetColumnIndex(ws, "TotalExperience")
    colAttr = GetColumnIndex(ws, "Attrition")
    
    If colExp = -1 Or colAttr = -1 Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim earlyAttr As Long, totalEarly As Long
    Dim i As Long
    
    For i = 2 To lastRow
        
        If ws.Cells(i, colExp).Value <= 2 Then
            
            totalEarly = totalEarly + 1
            
            If ws.Cells(i, colAttr).Value = "Yes" Then
                earlyAttr = earlyAttr + 1
            End If
            
        End If
        
    Next i
    
    Dim rate As Double
    If totalEarly > 0 Then rate = earlyAttr / totalEarly
    
    With ThisWorkbook.Sheets("Dashboard")
        .Range("B11").Value = rate
        .Range("B11").NumberFormat = "0.00%"
    End With

End Sub



Sub PromotionGapAnalysis()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim colPromo As Long
    colPromo = GetColumnIndex(ws, "LastPromotionDate")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    Dim highRisk As Long
    
    For i = 2 To lastRow
        
        If IsDate(ws.Cells(i, colPromo).Value) Then
            
            If DateDiff("yyyy", ws.Cells(i, colPromo).Value, Date) >= 3 Then
                highRisk = highRisk + 1
            End If
            
        End If
        
    Next i
    
    ' Output
    With ThisWorkbook.Sheets("Dashboard")
        .Range("B10").Value = highRisk
    End With

End Sub






Sub EmployeeRiskScore()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Clean_Data")
    
    Dim colOT As Long, colSat As Long, colSal As Long, colPromo As Long
    
    colOT = GetColumnIndex(ws, "Overtime")
    colSat = GetColumnIndex(ws, "JobSatisfaction")
    colSal = GetColumnIndex(ws, "Salary")
    colPromo = GetColumnIndex(ws, "LastPromotionDate")
    
    ' Validate columns
    If colOT = -1 Or colSat = -1 Or colSal = -1 Or colPromo = -1 Then
        MsgBox "Required columns missing for Risk Score!", vbCritical
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' ? Check if RiskScore column already exists
    Dim colRisk As Long
    colRisk = GetColumnIndex(ws, "RiskScore")
    
    If colRisk = -1 Then
        colRisk = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, colRisk).Value = "RiskScore"
    End If
    
    Dim i As Long
    Dim score As Integer
    
    For i = 2 To lastRow
        
        score = 0
        
        ' Overtime risk
        If ws.Cells(i, colOT).Value = "Yes" Then score = score + 2
        
        ' Low satisfaction risk
        If IsNumeric(ws.Cells(i, colSat).Value) Then
            If ws.Cells(i, colSat).Value <= 2 Then score = score + 2
        End If
        
        ' Low salary risk
        If IsNumeric(ws.Cells(i, colSal).Value) Then
            If ws.Cells(i, colSal).Value < 50000 Then score = score + 1
        End If
        
        ' Promotion gap risk
        If IsDate(ws.Cells(i, colPromo).Value) Then
            If DateDiff("yyyy", ws.Cells(i, colPromo).Value, Date) >= 3 Then
                score = score + 2
            End If
        End If
        
        ws.Cells(i, colRisk).Value = score
        
    Next i

    MsgBox "Employee Risk Scoring Completed", vbInformation

End Sub



