Sub RunFullPipeline()

    Call SetupDashboard
    Call RefreshData
    Call ValidateColumns
    
    Call CalculateAttrition
    Call DepartmentRisk
    Call GenerateInsights
    
    Call SalaryAttritionAnalysis
    Call ExperienceRisk
    Call EmployeeRiskScore
    Call CreateDeptChart
    
    MsgBox "HR Analytics Dashboard Ready ??", vbInformation

End Sub
