Option Explicit

Sub CalculateAverageSalary()
    Dim wsData As Worksheet
    Dim wsDash As Worksheet
    Dim lastRow As Long
    Dim avgSalary As Double

    Set wsData = ThisWorkbook.Sheets("HR_Employee_Data")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    lastRow = wsData.Cells(wsData.Rows.Count, "I").End(xlUp).Row

    avgSalary = Application.WorksheetFunction.Average(wsData.Range("I2:I" & lastRow))

    wsDash.Range("B8").Value = avgSalary
    wsDash.Range("B8").NumberFormat = "£#,##0.00"

    MsgBox "Average Salary Updated", vbInformation
End Sub