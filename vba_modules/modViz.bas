Sub CreateDeptChart()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    Dim chartObj As ChartObject
    
    ' Delete old chart
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Create new chart
    Set chartObj = ws.ChartObjects.Add(Left:=400, Top:=50, Width:=400, Height:=250)
    
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("D2:E8")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Attrition by Department"
    End With

End Sub
