Option Explicit

Sub CleanHRData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HR_Employee_Data")

    Application.ScreenUpdating = False

    ' Remove duplicate employee IDs
    ws.Range("A:T").RemoveDuplicates Columns:=1, Header:=xlYes

    ' Replace blanks with N/A
    ws.Cells.Replace What:="", Replacement:="N/A", LookAt:=xlWhole

    ' Standardize salary and bonus formatting
    ws.Columns("I:J").NumberFormat = "£#,##0.00"

    ' Format dates
    ws.Columns("H:H").NumberFormat = "dd-mmm-yyyy"

    ' Autofit all columns
    ws.Cells.EntireColumn.AutoFit

    Application.ScreenUpdating = True

    MsgBox "HR Data Cleaned Successfully", vbInformation
End Sub