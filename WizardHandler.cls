VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WizardHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
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

'wh.catch wizard_workbook
'
'np.TextBoxCW = wh.get_cw()
'np.TextBoxFaza = wh.get_faza()
'np.TextBoxPlt = wh.get_plt()
'np.TextBoxProj = wh.get_proj()

Private wiz As Workbook

Private details As Worksheet
Private master As Worksheet
Private puses As Worksheet


Public Sub catch(ByRef w As Workbook)
    Set wiz = w
    
    
    ' main sheets from the wizard - this is just a beginning
    Set details = wiz.Sheets("DETAILS")
    Set master = wiz.Sheets("MASTER")
    Set puses = wiz.Sheets("PICKUPS")
    
    
    
    Dim r As Range
    Set r = master.Range("A1")
    Set r = master.Range(r, r.End(xlDown).Offset(0, 50))
    
    ' Debug.Print r.Address
    
    Dim f As Range
    Set f = r.Find("*#*")
    
    If Not f Is Nothing Then
        MsgBox "W arkuszu master znjaduja sie hashe, makro przestaje dzialac, a Ty wez popraw dane!"
        End
    End If

    
    ' ------------------------------------------------------------------------------
End Sub


Public Sub go_with_6p_time()
    
    inner_6p_time master, details, puses
End Sub


Public Function get_cw() As String
    
    get_cw = ""
    
    ' get curr cw
    str_cw = CStr(Application.WorksheetFunction.IsoWeekNum(Date))
    
    If Len(str_cw) = 1 Then
        str_cw = "0" & str_cw
    End If
    get_cw = "" & Year(Date) & str_cw
End Function

Public Function get_plt() As String
    
    get_plt = CStr(details.Range("plant"))
End Function

Public Function get_plt_from_buffer() As String
    
    get_plt_from_buffer = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("O1"))
    
End Function

Public Function get_faza() As String
    get_faza = CStr(details.Range("faza"))
End Function

Public Function get_faza_from_buffer() As String
    get_faza_from_buffer = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("S1"))
End Function

Public Function get_proj() As String
    get_proj = CStr(details.Range("project"))
End Function

Public Function get_proj_from_buffer() As String
    get_proj_from_buffer = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("P1"))
End Function

Public Function get_my() As String
    get_my = CStr(details.Range("my"))
End Function

Public Function get_my_from_buffer() As String
    get_my_from_buffer = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("R1"))
End Function


Public Function get_biw_ga() As String
    get_biw_ga = CStr(details.Range("biw_ga"))
End Function

Public Function get_biw_ga_from_buffer() As String
    get_biw_ga_from_buffer = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("Q1"))
End Function
