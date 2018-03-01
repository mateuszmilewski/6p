VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form3Kolory 
   Caption         =   "3 KOLORY"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15720
   OleObjectBlob   =   "Form3Kolory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form3Kolory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnReset_Click()
    reset_3_kolory_form
End Sub

Private Sub BtnSubmit_Click()
    MsgBox "not implemented yet!"
    Hide
End Sub

Private Sub ComboBoxLink_Change()

    Dim l As T_Link, lr As Linker
    Set lr = New Linker
    

    If Trim(Me.ComboBoxLink.Value) <> "" Then
    
        If Not Trim(Me.ComboBoxLink.Value) Like "*,,*" Then
            ' to znaczy ze nie ma dziwnie ulozonych danych
            
            Set l = New T_Link
            Dim arr() As String
            arr = Split(Trim(Me.ComboBoxLink.Value), ",")
            
            l.zrob_mnie_z_argsow Trim(arr(0)), Trim(arr(1)), Trim(arr(2)), Trim(arr(3))
            '
            ' dobra znajdz mnie teraz w arkuszu del conf
            Dim dcsh As Worksheet
            Set dcsh = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
            
            Dim r As Range
            Set r = l.znajdz_siebie_w_arkuszu(dcsh)
            
            ' jestesmy na odpowiedniej "wysokosci" w arkuszu del conf
            Me.TextBoxEDI = CStr(r.Offset(0, SIXP.e_del_conf_edi - 1).Value)
            Me.TextBoxHO = CStr(r.Offset(0, SIXP.e_del_conf_ho - 1).Value)
            Me.TextBoxNA = CStr(r.Offset(0, SIXP.e_del_conf_na - 1).Value)
            Me.TextBoxOnStock = CStr(r.Offset(0, SIXP.e_del_conf_on_stock - 1).Value)
            
            Me.TextBoxForMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_mrd - 1).Value)
            Me.TextBoxAfterMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_mrd - 1).Value)
            
            Me.TextBoxFORSMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_smrd - 1).Value)
            Me.TextBoxAfterSMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_smrd - 1).Value)
            
            Me.TextBoxFORTMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_twomrd - 1).Value)
            Me.TextBoxAfterTMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_twomrd - 1).Value)
            
            Me.TextBoxFORSTMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_twosmrd - 1).Value)
            Me.TextBoxAfterSTMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_twosmrd - 1).Value)
            
            ' red
            Me.TextBoxOPEN = CStr(r.Offset(0, SIXP.e_del_conf_open - 1).Value)
            Me.TextBoxTooLate = CStr(r.Offset(0, SIXP.e_del_conf_too_late - 1).Value)
            Me.TextBoxPotITDC = CStr(r.Offset(0, SIXP.e_del_conf_pot_itdc - 1).Value)
            
            
            
            
            
        Else
            MsgBox "te dane sa jakies niefrasobliwe - nic z tym nie zrobie!"
        End If
    
    End If
End Sub
