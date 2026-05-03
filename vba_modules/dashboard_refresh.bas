Option Explicit

Sub RefreshDashboard()
    Application.ScreenUpdating = False

    ' Refresh all pivot tables and data connections
    ThisWorkbook.RefreshAll

    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Dashboard Refreshed Successfully", vbInformation
End Sub