Attribute VB_Name = "ResetModule"
Public Sub reset_6P(ictrl As IRibbonControl)
    
    Application.EnableEvents = True
    Application.CalculateFullRebuild
    Application.Calculation = xlCalculationAutomatic
    
End Sub
