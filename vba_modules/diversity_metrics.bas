Option Explicit

Sub UpdateGenderDiversity()
    Dim wsData As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim maleCount As Long
    Dim femaleCount As Long

    Set wsData = ThisWorkbook.Sheets("HR_Employee_Data")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    lastRow = wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row

    maleCount = Application.WorksheetFunction.CountIf(wsData.Range("D2:D" & lastRow), "Male")
    femaleCount = Application.WorksheetFunction.CountIf(wsData.Range("D2:D" & lastRow), "Female")

    wsDash.Range("B10").Value = maleCount
    wsDash.Range("B11").Value = femaleCount

    MsgBox "Gender Diversity Metrics Updated", vbInformation
End Sub