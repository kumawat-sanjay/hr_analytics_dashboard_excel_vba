Option Explicit

Sub ExportDashboardPDF()
    Dim filePath As String

    filePath = ThisWorkbook.Path & "\\HR_Analytics_Dashboard_Report.pdf"

    Sheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    MsgBox "Dashboard Exported to PDF Successfully", vbInformation
End Sub