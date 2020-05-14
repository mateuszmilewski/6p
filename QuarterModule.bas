Attribute VB_Name = "QuarterModule"
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

Public Sub uruchomLogikePrzechwytywaniaDanychZStaregoWybranegoQuartera(sh As Worksheet, wrknm As String)


    Application.ScreenUpdating = False

    SIXP.LoadingFormModule.showLoadingForm
    

    Dim qh As QuarterHandler
    Set qh = New QuarterHandler
    qh.setWrkNm wrknm
    qh.fillDictionaryWithTLinks sh
    qh.openFormWithDataFromQuarter
    Set qh = Nothing
    
    
    SIXP.LoadingFormModule.hideLoadingForm
    
    Application.ScreenUpdating = True
    
    ' MsgBox "import Quarter: ready!"
    
End Sub

Public Sub import_from_quarter(ictrl As IRibbonControl)


    Dim w As Workbook
    FormCatchQuarter.ListBoxFiles.Clear
    For Each w In Workbooks
        FormCatchQuarter.ListBoxFiles.addItem w.name
    Next w
    
    FormCatchQuarter.Show vbModeless
    
    
End Sub


