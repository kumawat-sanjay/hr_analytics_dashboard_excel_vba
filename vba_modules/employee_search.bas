Option Explicit

Sub SearchEmployeeByID()
    Dim empID As String
    Dim foundCell As Range
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("HR_Employee_Data")

    empID = InputBox("Enter Employee ID (Example: EMP00025)")

    If empID = "" Then Exit Sub

    Set foundCell = ws.Range("A:A").Find(What:=empID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ws.Activate
        foundCell.Select
        MsgBox "Employee Found: " & foundCell.Offset(0, 1).Value & " " & foundCell.Offset(0, 2).Value, vbInformation
    Else
        MsgBox "Employee ID Not Found", vbExclamation
    End If
End Sub