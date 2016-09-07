VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDelConfStatus 
   Caption         =   "FormDelConfStatus"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9465
   OleObjectBlob   =   "FormDelConfStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDelConfStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnSubmit_Click()

    ' text na guziki
    ' Global Const G_BTN_TEXT_ADD = "Dodaj"
    ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
    Hide
    inner_calc
    
    run_FormMain Me.LabelTitle
End Sub

Private Sub change_col_L_in_MAIN_worksheet(ByRef r As Range)
    
    ' tutaj sekcja, gdy dane juz zostaly dodane do arkusza order releases
    ' teraz nalezy odpowiednio o tym poinformowac arkusz glowny
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
    
        ' szukamy teraz w main
        ' ===================================================
        Dim rr As Range
        Set rr = ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Cells(1, 1)
        Do
            If CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(CStr(r.Offset(0, 3)))) = _
                CStr(Trim(rr) & ", " & Trim(rr.Offset(0, 1)) & ", " & Trim(rr.Offset(0, 2)) & ", " & Trim(CStr(rr.Offset(0, 3)))) Then
                    ' jest to samo w order release status sheet oraz to samo w main sheet
                    ' --------------------------------------------------------------------
                    ''
                    '
                    rr.Offset(0, SIXP.e_main_last_update_on_del_conf - 1) = Trim(CStr(rr.Offset(0, 3)))
                    '
                    ''
                    ' --------------------------------------------------------------------
                    Exit Do
            End If
            Set rr = rr.Offset(1, 0)
        Loop Until Trim(rr) = ""
        
        
        ' ===================================================
    
    
    
    
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
End Sub

Private Sub inner_calc()


   ' Public Enum E_DEL_CONF_ORDER
   '     e_del_conf_on_stock = 5
   '     e_del_conf_edi
   '     e_del_conf_ho
   '     e_del_conf_na
   '
   '
   '     e_del_conf_for_mrd
   '     e_del_conf_after_mrd
   '
   '     e_del_conf_for_smrd
   '     e_del_conf_after_smrd
   '
   '     e_del_conf_for_twomrd
   '     e_del_conf_after_twomrd
   '
   '     e_del_conf_for_twosmrd
   '     e_del_conf_after_twosmrd
   '
   '     e_del_conf_for_alt
   '     e_del_conf_after_alt
   '
   '     e_del_conf_open
   '     e_del_conf_pot_itdc
   '     e_del_conf_undef
   ' End Enum


    Dim r As Range
    
    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = arr(x)
        Next x
        
        
        give_data_to_ranges r
        change_col_L_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_L_in_MAIN_worksheet r
                    Exit Do
            End If
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        ' ===================================================
    Else
        MsgBox "fatal error on submitting!"
        End
    End If
End Sub

Private Sub give_data_to_ranges(ByRef r As Range)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_alt) = CStr(Me.TextBoxAfterALTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_mrd) = CStr(Me.TextBoxAfterMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_smrd) = CStr(Me.TextBoxAfterSMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_twomrd) = CStr(Me.TextBoxAfterTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_twosmrd) = CStr(Me.TextBoxAfterTSMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_edi) = CStr(Me.TextBoxEDI)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_alt) = CStr(Me.TextBoxForALTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_mrd) = CStr(Me.TextBoxForMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_smrd) = CStr(Me.TextBoxForSMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_twomrd) = CStr(Me.TextBoxForTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_twosmrd) = CStr(Me.TextBoxForTSMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_ho) = CStr(Me.TextBoxHO)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_na) = CStr(Me.TextBoxNA)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_on_stock) = CStr(Me.TextBoxOnStock)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_open) = CStr(Me.TextBoxOpen)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_pot_itdc) = CStr(Me.TextBoxPotITDC)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_undef) = CStr(Me.TextBoxUNDEF)
End Sub

' textboxes with qtyies bedzie w sumie 22 procedury wiec bierz sie do roboty
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------

'Private Sub NoOfVehLess_Click()
'    If IsNumeric(Me.TextBoxNoOfVeh) Then
'        If CLng(Me.TextBoxNoOfVeh) > 0 Then
'            tmp = CLng(Me.TextBoxNoOfVeh)
'            tmp = tmp - 1
'            Me.TextBoxNoOfVeh = CStr(tmp)
'        End If
'    End If
'End Sub

'Private Sub NoOfVehMore_Click()
'    If IsNumeric(Me.TextBoxNoOfVeh) Then
'        tmp = CLng(Me.TextBoxNoOfVeh)
'        tmp = tmp + 1
'        Me.TextBoxNoOfVeh = CStr(tmp)
'    End If
'End Sub

Private Sub AfterALTMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterALTMRD) Then
        If CLng(Me.TextBoxAfterALTMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterALTMRD)
            tmp = tmp - 1
            Me.TextBoxAfterALTMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub AfterALTMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterALTMRD) Then
        tmp = CLng(Me.TextBoxAfterALTMRD)
        tmp = tmp + 1
        Me.TextBoxAfterALTMRD = CStr(tmp)
    End If
End Sub

Private Sub AfterMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterMRD) Then
        If CLng(Me.TextBoxAfterMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterMRD)
            tmp = tmp - 1
            Me.TextBoxAfterMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub AfterMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterMRD) Then
        tmp = CLng(Me.TextBoxAfterMRD)
        tmp = tmp + 1
        Me.TextBoxAfterMRD = CStr(tmp)
    End If
End Sub

Private Sub AfterSMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterSMRD) Then
        If CLng(Me.TextBoxAfterSMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterSMRD)
            tmp = tmp - 1
            Me.TextBoxAfterSMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub AfterSMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterSMRD) Then
        tmp = CLng(Me.TextBoxAfterSMRD)
        tmp = tmp + 1
        Me.TextBoxAfterSMRD = CStr(tmp)
    End If
End Sub

Private Sub AfterTMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterTMRD) Then
        If CLng(Me.TextBoxAfterTMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterTMRD)
            tmp = tmp - 1
            Me.TextBoxAfterTMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub AfterTMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterTMRD) Then
        tmp = CLng(Me.TextBoxAfterTMRD)
        tmp = tmp + 1
        Me.TextBoxAfterTMRD = CStr(tmp)
    End If
End Sub

Private Sub AfterTSMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterTSMRD) Then
        If CLng(Me.TextBoxAfterTSMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterTSMRD)
            tmp = tmp - 1
            Me.TextBoxAfterTSMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub AfterTSMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterTSMRD) Then
        tmp = CLng(Me.TextBoxAfterTSMRD)
        tmp = tmp + 1
        Me.TextBoxAfterTSMRD = CStr(tmp)
    End If
End Sub

Private Sub EDILess_Click()
    If IsNumeric(Me.TextBoxEDI) Then
        If CLng(Me.TextBoxEDI) > 0 Then
            tmp = CLng(Me.TextBoxEDI)
            tmp = tmp - 1
            Me.TextBoxEDI = CStr(tmp)
        End If
    End If
End Sub

Private Sub EDIMore_Click()
    If IsNumeric(Me.TextBoxEDI) Then
        tmp = CLng(Me.TextBoxEDI)
        tmp = tmp + 1
        Me.TextBoxEDI = CStr(tmp)
    End If
End Sub

Private Sub ForALTMRDLess_Click()
    If IsNumeric(Me.TextBoxForALTMRD) Then
        If CLng(Me.TextBoxForALTMRD) > 0 Then
            tmp = CLng(Me.TextBoxForALTMRD)
            tmp = tmp - 1
            Me.TextBoxForALTMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub ForALTMRDMore_Click()
    If IsNumeric(Me.TextBoxForALTMRD) Then
        tmp = CLng(Me.TextBoxForALTMRD)
        tmp = tmp + 1
        Me.TextBoxForALTMRD = CStr(tmp)
    End If
End Sub

Private Sub ForMRDLess_Click()
    If IsNumeric(Me.TextBoxForMRD) Then
        If CLng(Me.TextBoxForMRD) > 0 Then
            tmp = CLng(Me.TextBoxForMRD)
            tmp = tmp - 1
            Me.TextBoxForMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub ForMRDMore_Click()
    If IsNumeric(Me.TextBoxForMRD) Then
        tmp = CLng(Me.TextBoxForMRD)
        tmp = tmp + 1
        Me.TextBoxForMRD = CStr(tmp)
    End If
End Sub

Private Sub ForSMRDLess_Click()
    If IsNumeric(Me.TextBoxForSMRD) Then
        If CLng(Me.TextBoxForSMRD) > 0 Then
            tmp = CLng(Me.TextBoxForSMRD)
            tmp = tmp - 1
            Me.TextBoxForSMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub ForSMRDMore_Click()
    If IsNumeric(Me.TextBoxForSMRD) Then
        tmp = CLng(Me.TextBoxForSMRD)
        tmp = tmp + 1
        Me.TextBoxForSMRD = CStr(tmp)
    End If
End Sub

Private Sub ForTMRDLess_Click()
    If IsNumeric(Me.TextBoxForTMRD) Then
        If CLng(Me.TextBoxForTMRD) > 0 Then
            tmp = CLng(Me.TextBoxForTMRD)
            tmp = tmp - 1
            Me.TextBoxForTMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub ForTMRDMore_Click()
    If IsNumeric(Me.TextBoxForTMRD) Then
        tmp = CLng(Me.TextBoxForTMRD)
        tmp = tmp + 1
        Me.TextBoxForTMRD = CStr(tmp)
    End If
End Sub

Private Sub ForTSMRDLess_Click()
    If IsNumeric(Me.TextBoxForTSMRD) Then
        If CLng(Me.TextBoxForTSMRD) > 0 Then
            tmp = CLng(Me.TextBoxForTSMRD)
            tmp = tmp - 1
            Me.TextBoxForTSMRD = CStr(tmp)
        End If
    End If
End Sub

Private Sub ForTSMRDMore_Click()
    If IsNumeric(Me.TextBoxForTSMRD) Then
        tmp = CLng(Me.TextBoxForTSMRD)
        tmp = tmp + 1
        Me.TextBoxForTSMRD = CStr(tmp)
    End If
End Sub

Private Sub HOLess_Click()
    If IsNumeric(Me.TextBoxHO) Then
        If CLng(Me.TextBoxHO) > 0 Then
            tmp = CLng(Me.TextBoxHO)
            tmp = tmp - 1
            Me.TextBoxHO = CStr(tmp)
        End If
    End If
End Sub

Private Sub HOMore_Click()
    If IsNumeric(Me.TextBoxHO) Then
        tmp = CLng(Me.TextBoxHO)
        tmp = tmp + 1
        Me.TextBoxHO = CStr(tmp)
    End If
End Sub

Private Sub NALess_Click()
    If IsNumeric(Me.TextBoxNA) Then
        If CLng(Me.TextBoxNA) > 0 Then
            tmp = CLng(Me.TextBoxNA)
            tmp = tmp - 1
            Me.TextBoxNA = CStr(tmp)
        End If
    End If
End Sub

Private Sub NAMore_Click()
    If IsNumeric(Me.TextBoxNA) Then
        tmp = CLng(Me.TextBoxNA)
        tmp = tmp + 1
        Me.TextBoxNA = CStr(tmp)
    End If
End Sub

Private Sub OnStockLess_Click()
    If IsNumeric(Me.TextBoxOnStock) Then
        If CLng(Me.TextBoxOnStock) > 0 Then
            tmp = CLng(Me.TextBoxOnStock)
            tmp = tmp - 1
            Me.TextBoxOnStock = CStr(tmp)
        End If
    End If
End Sub

Private Sub OnStockMore_Click()
    If IsNumeric(Me.TextBoxOnStock) Then
        tmp = CLng(Me.TextBoxOnStock)
        tmp = tmp + 1
        Me.TextBoxOnStock = CStr(tmp)
    End If
End Sub

Private Sub OpenLess_Click()
    If IsNumeric(Me.TextBoxOpen) Then
        If CLng(Me.TextBoxOpen) > 0 Then
            tmp = CLng(Me.TextBoxOpen)
            tmp = tmp - 1
            Me.TextBoxOpen = CStr(tmp)
        End If
    End If
End Sub

Private Sub OpenMore_Click()
    If IsNumeric(Me.TextBoxOpen) Then
        tmp = CLng(Me.TextBoxOpen)
        tmp = tmp + 1
        Me.TextBoxOpen = CStr(tmp)
    End If
End Sub

Private Sub PotITDCLess_Click()
    If IsNumeric(Me.TextBoxPotITDC) Then
        If CLng(Me.TextBoxPotITDC) > 0 Then
            tmp = CLng(Me.TextBoxPotITDC)
            tmp = tmp - 1
            Me.TextBoxPotITDC = CStr(tmp)
        End If
    End If
End Sub

Private Sub PotITDCMore_Click()
    If IsNumeric(Me.TextBoxPotITDC) Then
        tmp = CLng(Me.TextBoxPotITDC)
        tmp = tmp + 1
        Me.TextBoxPotITDC = CStr(tmp)
    End If
End Sub

Private Sub UNDEFLess_Click()
    If IsNumeric(Me.TextBoxUNDEF) Then
        If CLng(Me.TextBoxUNDEF) > 0 Then
            tmp = CLng(Me.TextBoxUNDEF)
            tmp = tmp - 1
            Me.TextBoxUNDEF = CStr(tmp)
        End If
    End If
End Sub

Private Sub UNDEFMore_Click()
    If IsNumeric(Me.TextBoxUNDEF) Then
        tmp = CLng(Me.TextBoxUNDEF)
        tmp = tmp + 1
        Me.TextBoxUNDEF = CStr(tmp)
    End If
End Sub
