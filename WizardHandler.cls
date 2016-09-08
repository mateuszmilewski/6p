VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WizardHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    
    Set details = wiz.Sheets("DETAILS")
    Set master = wiz.Sheets("MASTER")
    Set puses = wiz.Sheets("PICKUPS")
End Sub


Public Function get_cw() As String
    
    get_cw = ""
    
    ' get curr cw
    str_cw = CStr(Application.WorksheetFunction.IsoWeekNum(Date))
    
    If Len(str_cw) = 1 Then
        str_cw = "0" & cstr_cw
    End If
    get_cw = "" & Year(Date) & str_cw
End Function

Public Function get_plt() As String
    
    get_plt = CStr(details.Range("plant"))
End Function

Public Function get_faza() As String
    get_faza = CStr(details.Range("faza"))
End Function

Public Function get_proj() As String
    get_proj = CStr(details.Range("project"))
End Function
