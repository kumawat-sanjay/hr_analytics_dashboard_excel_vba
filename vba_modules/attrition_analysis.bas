Option Explicit

Sub GenerateAttritionReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim attritionCount As Long
    Dim totalEmployees As Long
    Dim attritionRate As Double

    Set wsData = ThisWorkbook.Sheets("HR_Employee_Data")
    Set wsReport = ThisWorkbook.Sheets("Dashboard")

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    totalEmployees = lastRow - 1

    attritionCount = Application.WorksheetFunction.CountIf(wsData.Range("L2:L" & lastRow), "Yes")

    attritionRate = (attritionCount / totalEmployees) * 100

    wsReport.Range("B5").Value = totalEmployees
    wsReport.Range("B6").Value = attritionCount
    wsReport.Range("B7").Value = Round(attritionRate, 2) & "%"

    MsgBox "Attrition Report Updated", vbInformation
End Sub