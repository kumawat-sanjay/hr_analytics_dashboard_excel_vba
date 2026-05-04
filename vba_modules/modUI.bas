Sub SetupDashboard()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "HR ANALYTICS DASHBOARD"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1:E1").Merge
    ws.Range("A1").HorizontalAlignment = xlCenter
    
    ' KPI Labels
    ws.Range("A3").Value = "Total Employees"
    ws.Range("A4").Value = "Attrition Count"
    ws.Range("A5").Value = "Attrition Rate"
    
    ' Insight
    ws.Range("A7").Value = "Insight"
    
    ' Advanced Metrics
    ws.Range("A10").Value = "Low Salary Attrition"
    ws.Range("A11").Value = "Early Career Attrition"
    
    ' Department Table
    ws.Range("D2").Value = "Department"
    ws.Range("E2").Value = "Attrition Rate"
    ws.Range("D2:E2").Font.Bold = True
    
    ' Column widths
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 18
    
End Sub
