VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewProj 
   Caption         =   "Projekt"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   OleObjectBlob   =   "NewProj.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEdit_Click()
    
    Dim m As Worksheet, r As Range
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    
    If ThisWorkbook.ActiveSheet.Name = m.Name Then
    
    
        ' ==================================================
        Set r = validate_and_then_go_to_active_cell
        ' ==================================================
        
        If r Is Nothing Then
            MsgBox "Akcja nie jest dozwolona!"
        Else
            r.Value = Me.TextBoxProj
            r.Offset(0, 1).Value = Me.TextBoxPlt
            r.Offset(0, 2).Value = Me.TextBoxFaza
            r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
            r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
        End If
    Else
        ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
        MsgBox "nie mozna wykonac akcji w tej lokalizacji pliku - makro samo Cie przesunelo na wlasciwy arkusz."
    End If
End Sub

Private Sub BtnSubmit_Click()
    ' tutaj dodajemy nowy projekt na spod w arkuszu main
    
    Dim m As Worksheet
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    Dim r As Range
    Set r = validate_and_then_go_to_first_empty_cell(m)
    
    r.Value = Me.TextBoxProj
    r.Offset(0, 1).Value = Me.TextBoxPlt
    r.Offset(0, 2).Value = Me.TextBoxFaza
    r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
    r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
End Sub



Private Function validate_and_then_go_to_first_empty_cell(ByRef m As Worksheet) As Range
    
    Dim r As Range
    Set r = m.Cells(1, 1)
    Do
        If CStr(Me.TextBoxProj) = Trim(CStr(r)) Then
            If CStr(Me.TextBoxPlt) = Trim(CStr(r.Offset(0, 1))) Then
                If Trim(CStr(r.Offset(0, 2).Value)) = Trim(CStr(Me.TextBoxFaza)) Then
                    If Trim(CStr(r.Offset(0, 3).Value)) = CStr(Me.TextBoxCW) Then
                        MsgBox "duplikat! apka konczy dzialanie!"
                        End
                    End If
                End If
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Set validate_and_then_go_to_first_empty_cell = r
End Function

Private Function validate_and_then_go_to_active_cell() As Range
    
    Dim r As Range
    Set r = ActiveCell
    Set r = r.Parent.Cells(r.Row, 1)
    Do
        If Trim(CStr(r)) <> "" _
            Or Trim(CStr(r.Offset(0, 1))) <> "" _
            Or Trim(CStr(r.Offset(0, 2).Value)) <> "" _
            Or Trim(CStr(r.Offset(0, 3).Value)) <> "" Then
            
            Set validate_and_then_go_to_active_cell = r
            Exit Function
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Set validate_and_then_go_to_active_cell = Nothing
End Function

Private Sub ComboBoxFAZA_Change()
    Me.TextBoxFaza = CStr(Me.ComboBoxFAZA.Value)
End Sub

Private Sub ComboBoxPLT_Change()
    Me.TextBoxPlt = CStr(Me.ComboBoxPLT.Value)
End Sub



Private Sub DTPicker1_Change()
    Me.TextBoxCW = SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPicker1.Value))
End Sub
