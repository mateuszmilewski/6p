Attribute VB_Name = "GlobalCrossTriangleCircleModule"
Public Function putCross() As Range
    Set putCross = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("cross")
End Function
Public Function putTriangle() As Range
    Set putTriangle = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("triangle")
End Function
Public Function putCircle() As Range
    Set putCircle = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("circle")
End Function


Public Sub setCross(ictrl As IRibbonControl)
    ActiveCell = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("cross")
End Sub
Public Sub setTriangle(ictrl As IRibbonControl)
    ActiveCell = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("triangle")
End Sub
Public Sub setCircle(ictrl As IRibbonControl)
    ActiveCell = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("circle")
End Sub
