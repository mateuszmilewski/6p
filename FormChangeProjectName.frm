VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormChangeProjectName 
   Caption         =   "Change Project Name"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8115
   OleObjectBlob   =   "FormChangeProjectName.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormChangeProjectName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnChange_Click()

    Hide

    ' if zawiera funkcje sprawdzajaca z boolean jako wyjscie
    ' jednak sam boolean dyktowany bedzie yes/no od msgboxa od usera
    ' po zweryfikowaniu informacji o ilosci zmian dla potencjalnych projektow
    If sprawdzNajpierwJakiePotencjalnieZmianyZajda() Then
    
    
        Dim l As T_Link
        Set l = New T_Link
        l.zrob_mnie_z_argsow CStr(Me.TextBoxCurrProj), CStr(Me.TextBoxCurrPltCode), CStr(Me.TextBoxCurrFaza), CStr(Me.TextBoxCurrCw)
        
        
        
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_main_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_order_release_status_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_recent_build_plan_changes_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_cont_pnoc_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_osea_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_totals_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_resp_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_del_conf_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_open_issues_sh_nm)
        przeszukajArkuszIUsunWystapienia l, CStr(SIXP.G_xq_sh_nm)
    
    End If
End Sub


Private Function przeszukajArkuszIUsunWystapienia(l As T_Link, shnm As String)
    
    If shnm <> "" Then
        If Not l Is Nothing Then
            
            
            ' main logic for data deletion
            ' ---------------------------------------------
            
            Dim Sh As Worksheet, r As Range
            Set Sh = ThisWorkbook.Sheets(CStr(shnm))
            
            Set r = l.znajdz_siebie_w_arkuszu(Sh)
            
            If Not r Is Nothing Then
            
            
                zmienNazweRekorduZgodnieZNowymZapisem r
                przeszukajArkuszIUsunWystapienia l, shnm
            Else
                If r Is Nothing Then
                    Exit Function
                Else
                    MsgBox "zbior pusty - ten if nie moze sie wydarzyc - change project name!"
                End If
            End If
                
                
            
            ' ---------------------------------------------
            
        End If
    End If
End Function


Private Sub zmienNazweRekorduZgodnieZNowymZapisem(mr As Range)
    
    Dim r As Range
    Set r = mr.Parent.Cells(mr.Row, 1)
    
    r.Value = CStr(Me.TextBoxNewProj.Value)
    r.Offset(0, 1).Value = CStr(Me.TextBoxNewPltCode.Value)
    r.Offset(0, 2).Value = CStr(Me.TextBoxNewFaza.Value)
    r.Offset(0, 3).Value = CStr(Me.TextBoxNewCw.Value)
End Sub


Private Function sprawdzNajpierwJakiePotencjalnieZmianyZajda() As Boolean
    
    sprawdzNajpierwJakiePotencjalnieZmianyZajda = False
    
    
    trescMsgboxa = przygotujTrescMsgboxa()
    
    answer = MsgBox(trescMsgboxa, vbYesNo)
    
    
    If answer = vbYes Then
        sprawdzNajpierwJakiePotencjalnieZmianyZajda = True
    Else
    
    End If
End Function

Private Function przygotujTrescMsgboxa() As String
    
    przygotujTrescMsgboxa = _
        "Czy napewno chcesz zamienic na nowe nazwy z tego 6P wszystkie wystapienia: " & _
        CStr(Me.TextBoxCurrProj.Value) & _
        CStr(Me.TextBoxCurrPltCode.Value) & ", " & _
        CStr(Me.TextBoxCurrFaza.Value) & ", " & _
        CStr(Me.TextBoxCurrCw.Value) & "?"
End Function


Private Sub BtnClose_Click()
    Me.Hide
End Sub

Private Sub BtnCopy_Click()


    ' copy info from current textboxes to new textboxes
    
    With Me
        .TextBoxNewProj.Value = CStr(.TextBoxCurrProj.Value)
        .TextBoxNewPltCode.Value = CStr(.TextBoxCurrPltCode.Value)
        .TextBoxNewFaza.Value = CStr(.TextBoxCurrFaza.Value)
        .TextBoxNewCw.Value = CStr(.TextBoxCurrCw.Value)
    End With
End Sub
