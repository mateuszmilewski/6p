VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewProj 
   Caption         =   "Nowy Projekt"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "NewProj.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnEdit_Click()
    
    Dim m As Worksheet
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    Dim r As Range
    Set r = ActiveCell.Parent.Cells(ActiveCell.Row, 1)
    
    r.Value = Me.TextBoxProj
    r.Offset(0, 1).Value = Me.TextBoxPlt
    r.Offset(0, 2).Value = Me.TextBoxFaza
    r.Offset(0, 3).Value = Me.TextBoxCW
    r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
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
    r.Offset(0, 3).Value = Me.TextBoxCW
    r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
End Sub


Private Function got_to_this_project(ByRef m As Worksheet) As Range
    
    Dim r As Range
    Set r = m.Cells(1, 1)
    Do
        If CStr(Me.TextBoxProj) = CStr(r) Then
            If CStr(Me.TextBoxPlt) = CStr(r.Offset(0, 1)) Then
                If CStr(r.Offset(0, 2).Value) = CStr(Me.TextBoxFaza) Then
                    If CStr(r.Offset(0, 3).Value) = CStr(Me.TextBoxCW) Then
                        If CStr(r.Offset(0, 4).Value) = CStr(Me.ComboBoxStatus.Value) Then
                        
                            Set go_to_first_empty_cell = r
                        End If
                    End If
                End If
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
End Function

Private Function validate_and_then_go_to_first_empty_cell(ByRef m As Worksheet) As Range
    
    Dim r As Range
    Set r = m.Cells(1, 1)
    Do
        If CStr(Me.TextBoxProj) = CStr(r) Then
            If CStr(Me.TextBoxPlt) = CStr(r.Offset(0, 1)) Then
                If CStr(r.Offset(0, 2).Value) = CStr(Me.TextBoxFaza) Then
                    If CStr(r.Offset(0, 3).Value) = CStr(Me.TextBoxCW) Then
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
