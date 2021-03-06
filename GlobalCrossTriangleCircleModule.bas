Attribute VB_Name = "GlobalCrossTriangleCircleModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

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
