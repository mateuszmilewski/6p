VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDelConfStatus 
   Caption         =   "FormDelConfStatus"
   ClientHeight    =   10560
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








Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnSubmit_Click()

    ' text na guziki
    ' Global Const G_BTN_TEXT_ADD = "Dodaj"
    ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
    'Hide
    inner_calc
    
    ' run_FormMain Me.LabelTitle
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
   


    Dim r As Range
    
    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For X = 0 To 3
            r.Offset(0, X) = Trim(arr(X))
        Next X
        
        
        give_data_to_ranges r
        change_col_L_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
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
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_smrd) = CStr(Me.TextBoxFORSMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_twomrd) = CStr(Me.TextBoxFORTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_twosmrd) = CStr(Me.TextBoxForTSMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_ho) = CStr(Me.TextBoxHO)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_na) = CStr(Me.TextBoxNA)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_on_stock) = CStr(Me.TextBoxOnStock)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_open) = CStr(Me.TextBoxOPEN)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_too_late) = CStr(Me.TextBoxTooLate)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_pot_itdc) = CStr(Me.TextBoxPotITDC)
    
    ' new data - new del confs
    ' ---------------------------------------------------------------------------------------
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_alttwomrd) = CStr(Me.TextBoxForALTTWOMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_alttwomrd) = CStr(Me.TextBoxAfterALTTWOMRD)
    
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_salttwomrd) = CStr(Me.TextBoxForSALTTWOMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_salttwomrd) = CStr(Me.TextBoxAfterSALTTWOMRD)
    
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_oncostmrd) = CStr(Me.TextBoxForONCOSTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_oncostmrd) = CStr(Me.TextBoxAfterONCOSTMRD)
    
    r.Parent.Cells(r.Row, SIXP.e_del_conf_for_soncostmrd) = CStr(Me.TextBoxForSONCOSTMRD)
    r.Parent.Cells(r.Row, SIXP.e_del_conf_after_soncostmrd) = CStr(Me.TextBoxAfterSONCOSTMRD)
    ' ---------------------------------------------------------------------------------------
End Sub

' textboxes with qtyies bedzie w sumie 22 procedury wiec bierz sie do roboty
' w 2018 doszly kolejne 8 textboxow
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


Private Sub AfterSONCOSTMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterSONCOSTMRD) Then
        If CLng(Me.TextBoxAfterSONCOSTMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterSONCOSTMRD)
            tmp = tmp - 1
            Me.TextBoxAfterSONCOSTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterSONCOSTMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterSONCOSTMRD) Then
        tmp = CLng(Me.TextBoxAfterSONCOSTMRD)
        tmp = tmp + 1
        Me.TextBoxAfterSONCOSTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub


Private Sub ForSONCOSTMRDLess_Click()
    If IsNumeric(Me.TextBoxForSONCOSTMRD) Then
        If CLng(Me.TextBoxForSONCOSTMRD) > 0 Then
            tmp = CLng(Me.TextBoxForSONCOSTMRD)
            tmp = tmp - 1
            Me.TextBoxForSONCOSTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForSONCOSTMRDMore_Click()
    If IsNumeric(Me.TextBoxForSONCOSTMRD) Then
        tmp = CLng(Me.TextBoxForSONCOSTMRD)
        tmp = tmp + 1
        Me.TextBoxForSONCOSTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterONCOSTMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterONCOSTMRD) Then
        If CLng(Me.TextBoxAfterONCOSTMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterONCOSTMRD)
            tmp = tmp - 1
            Me.TextBoxAfterONCOSTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterONCOSTMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterONCOSTMRD) Then
        tmp = CLng(Me.TextBoxAfterONCOSTMRD)
        tmp = tmp + 1
        Me.TextBoxAfterONCOSTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForONCOSTMRDLess_Click()
    If IsNumeric(Me.TextBoxForONCOSTMRD) Then
        If CLng(Me.TextBoxForONCOSTMRD) > 0 Then
            tmp = CLng(Me.TextBoxForONCOSTMRD)
            tmp = tmp - 1
            Me.TextBoxForONCOSTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForONCOSTMRDMore_Click()
    If IsNumeric(Me.TextBoxForONCOSTMRD) Then
        tmp = CLng(Me.TextBoxForONCOSTMRD)
        tmp = tmp + 1
        Me.TextBoxForONCOSTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub


Private Sub AfterSALTTWOMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterSALTTWOMRD) Then
        If CLng(Me.TextBoxAfterSALTTWOMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterSALTTWOMRD)
            tmp = tmp - 1
            Me.TextBoxAfterSALTTWOMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterSALTTWOMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterSALTTWOMRD) Then
        tmp = CLng(Me.TextBoxAfterSALTTWOMRD)
        tmp = tmp + 1
        Me.TextBoxAfterSALTTWOMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach

End Sub

Private Sub ForSALTTWOMRDLess_Click()

    If IsNumeric(Me.TextBoxForSALTTWOMRD) Then
        If CLng(Me.TextBoxForSALTTWOMRD) > 0 Then
            tmp = CLng(Me.TextBoxForSALTTWOMRD)
            tmp = tmp - 1
            Me.TextBoxForSALTTWOMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach

End Sub

Private Sub ForSALTTWOMRDMore_Click()
    If IsNumeric(Me.TextBoxForSALTTWOMRD) Then
        tmp = CLng(Me.TextBoxForSALTTWOMRD)
        tmp = tmp + 1
        Me.TextBoxForSALTTWOMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub


Private Sub ForALTTWOMRDLess_Click()

    If IsNumeric(Me.TextBoxForALTTWOMRD) Then
        If CLng(Me.TextBoxForALTTWOMRD) > 0 Then
            tmp = CLng(Me.TextBoxForALTTWOMRD)
            tmp = tmp - 1
            Me.TextBoxForALTTWOMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach

End Sub

Private Sub ForALTTWOMRDMore_Click()
    If IsNumeric(Me.TextBoxForALTTWOMRD) Then
        tmp = CLng(Me.TextBoxForALTTWOMRD)
        tmp = tmp + 1
        Me.TextBoxForALTTWOMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub


Private Sub AfterALTTWOMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterALTTWOMRD) Then
        If CLng(Me.TextBoxAfterALTTWOMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterALTTWOMRD)
            tmp = tmp - 1
            Me.TextBoxAfterALTTWOMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterALTTWOMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterALTTWOMRD) Then
        tmp = CLng(Me.TextBoxAfterALTTWOMRD)
        tmp = tmp + 1
        Me.TextBoxAfterALTTWOMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterALTMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterALTMRD) Then
        If CLng(Me.TextBoxAfterALTMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterALTMRD)
            tmp = tmp - 1
            Me.TextBoxAfterALTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterALTMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterALTMRD) Then
        tmp = CLng(Me.TextBoxAfterALTMRD)
        tmp = tmp + 1
        Me.TextBoxAfterALTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterMRD) Then
        If CLng(Me.TextBoxAfterMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterMRD)
            tmp = tmp - 1
            Me.TextBoxAfterMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterMRD) Then
        tmp = CLng(Me.TextBoxAfterMRD)
        tmp = tmp + 1
        Me.TextBoxAfterMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterSMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterSMRD) Then
        If CLng(Me.TextBoxAfterSMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterSMRD)
            tmp = tmp - 1
            Me.TextBoxAfterSMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterSMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterSMRD) Then
        tmp = CLng(Me.TextBoxAfterSMRD)
        tmp = tmp + 1
        Me.TextBoxAfterSMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterTMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterTMRD) Then
        If CLng(Me.TextBoxAfterTMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterTMRD)
            tmp = tmp - 1
            Me.TextBoxAfterTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterTMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterTMRD) Then
        tmp = CLng(Me.TextBoxAfterTMRD)
        tmp = tmp + 1
        Me.TextBoxAfterTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterTSMRDLess_Click()
    If IsNumeric(Me.TextBoxAfterTSMRD) Then
        If CLng(Me.TextBoxAfterTSMRD) > 0 Then
            tmp = CLng(Me.TextBoxAfterTSMRD)
            tmp = tmp - 1
            Me.TextBoxAfterTSMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub AfterTSMRDMore_Click()
    If IsNumeric(Me.TextBoxAfterTSMRD) Then
        tmp = CLng(Me.TextBoxAfterTSMRD)
        tmp = tmp + 1
        Me.TextBoxAfterTSMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub DelConfFromBuffBtn_Click()
    
    
    
    
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    
    ' rngv - range values, range labels
    Dim rngv As Range, rngl As Range, h1_total
    Set rngv = buff.Cells(17, 1)
    Set rngl = buff.Cells(16, 1)
    Set h1_total = buff.Range("H1")
    
    
    With buff
    
        ' 15: BEFORE OR ON/AFTER MRD
        ' 16: BEFORE/AFTER MRD - labels all
        ' 17: values
        Me.TextBoxAfterALTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_MRD))
        Me.TextBoxForALTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_MRD))
        
        Me.TextBoxAfterMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_MRD))
        Me.TextBoxForMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_MRD))
        
        Me.TextBoxAfterSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_MRD))
        Me.TextBoxFORSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_MRD))
        
        ' obsolete
        'Me.TextBoxAfterTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_TWO_MRD))
        'Me.TextBoxFORTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_TWO_MRD))
        
        'Me.TextBoxAfterTSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_TWO_Staggered_MRD))
        'Me.TextBoxForTSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_TWO_Staggered_MRD))
        
        
        ' new
        ' ---------------------------------------------------------------------------------------------------------------------
        
        Me.TextBoxAfterALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_TWO_MRD))
        Me.TextBoxForALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_TWO_MRD))
        
        Me.TextBoxAfterSALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        Me.TextBoxForSALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        
        Me.TextBoxAfterONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ONCOST_MRD))
        Me.TextBoxForONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ONCOST_MRD))
        
        Me.TextBoxAfterSONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_ONCOST_MRD))
        Me.TextBoxForSONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_ONCOST_MRD))
        
        ' ---------------------------------------------------------------------------------------------------------------------
    End With
        
        
        
    
    ' DEL CONF, WHICH IS NOT CONNECTED WITH MRD PARAM.
    
    Set rngv = buff.Cells(13, 1)
    Set rngl = buff.Cells(12, 1)
    
    With buff
        
        
        ' greens
        Me.TextBoxOnStock = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_ON_STOCK))
        Me.TextBoxEDI = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_EDI))
        Me.TextBoxHO = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_HO))
        Me.TextBoxNA = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_NA))
        
        
        'reds
        ' jednak tutaj jest powazny problem poniewaz wizard jako tako nie bierze pod uwagi podzialu NOKow
        ' traktowane sa one normlanie jako blanki bez wiekszego zglebiania
        ' zatem ponizsza logika nie ma sensu zeby byla powielana w wykorzystaniu tak jak to mialo miejsce w greensach
        ' czy polach uzaleznionych od MRD
        ' me.TextBoxOpen = cstr(sixp.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl...)
        Me.TextBoxOPEN = "0"
        Me.TextBoxTooLate = "0"
        ' wyjatekiem jest pot itdc, poniewaz jako tako mozna wyrazic ten element za pomoca stringa zamieszcznego w wizardzie
        Me.TextBoxPotITDC = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_POTITDC))
    End With
    
    sprawdz_czy_pola_przypadkiem_nie_sa_puste
    
    Dim suma_wszystkich_boxow As Long
    
    ' TWO & STWO removed
    suma_wszystkich_boxow = CLng(Me.TextBoxAfterALTMRD.Value) + _
        CLng(Me.TextBoxAfterMRD.Value) + CLng(Me.TextBoxAfterSMRD.Value) + _
        CLng(Me.TextBoxEDI.Value) + CLng(Me.TextBoxForALTMRD.Value) + _
        CLng(Me.TextBoxForMRD.Value) + CLng(Me.TextBoxFORSMRD.Value) + _
        CLng(Me.TextBoxFORTMRD.Value) + CLng(Me.TextBoxForTSMRD.Value) + _
        CLng(Me.TextBoxHO.Value) + CLng(Me.TextBoxNA.Value) + CLng(Me.TextBoxOnStock.Value) + _
        CLng(Me.TextBoxPotITDC.Value) + CLng(Me.TextBoxTooLate.Value)
        
    ' now add new fields
    suma_wszystkich_boxow = suma_wszystkich_boxow + _
        CLng(Me.TextBoxAfterALTTWOMRD.Value) + CLng(Me.TextBoxForALTTWOMRD.Value) + _
        CLng(Me.TextBoxAfterSALTTWOMRD.Value) + CLng(Me.TextBoxForALTTWOMRD.Value) + _
        CLng(Me.TextBoxAfterONCOSTMRD.Value) + CLng(Me.TextBoxForONCOSTMRD.Value) + _
        CLng(Me.TextBoxAfterSONCOSTMRD.Value) + CLng(Me.TextBoxForSONCOSTMRD.Value)
        
    Me.TextBoxOPEN = CLng(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("h1").Value) - suma_wszystkich_boxow
    
    zmiany_na_totalach
End Sub

Private Sub EDILess_Click()
    If IsNumeric(Me.TextBoxEDI) Then
        If CLng(Me.TextBoxEDI) > 0 Then
            tmp = CLng(Me.TextBoxEDI)
            tmp = tmp - 1
            Me.TextBoxEDI = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub EDIMore_Click()
    If IsNumeric(Me.TextBoxEDI) Then
        tmp = CLng(Me.TextBoxEDI)
        tmp = tmp + 1
        Me.TextBoxEDI = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForALTMRDLess_Click()
    If IsNumeric(Me.TextBoxForALTMRD) Then
        If CLng(Me.TextBoxForALTMRD) > 0 Then
            tmp = CLng(Me.TextBoxForALTMRD)
            tmp = tmp - 1
            Me.TextBoxForALTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForALTMRDMore_Click()
    If IsNumeric(Me.TextBoxForALTMRD) Then
        tmp = CLng(Me.TextBoxForALTMRD)
        tmp = tmp + 1
        Me.TextBoxForALTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub



Private Sub ForMRDLess_Click()
    If IsNumeric(Me.TextBoxForMRD) Then
        If CLng(Me.TextBoxForMRD) > 0 Then
            tmp = CLng(Me.TextBoxForMRD)
            tmp = tmp - 1
            Me.TextBoxForMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForMRDMore_Click()
    If IsNumeric(Me.TextBoxForMRD) Then
        tmp = CLng(Me.TextBoxForMRD)
        tmp = tmp + 1
        Me.TextBoxForMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub



Private Sub ForSMRDLess_Click()
    If IsNumeric(Me.TextBoxFORSMRD) Then
        If CLng(Me.TextBoxFORSMRD) > 0 Then
            tmp = CLng(Me.TextBoxFORSMRD)
            tmp = tmp - 1
            Me.TextBoxFORSMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForSMRDMore_Click()
    If IsNumeric(Me.TextBoxFORSMRD) Then
        tmp = CLng(Me.TextBoxFORSMRD)
        tmp = tmp + 1
        Me.TextBoxFORSMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub



Private Sub ForTMRDLess_Click()
    If IsNumeric(Me.TextBoxFORTMRD) Then
        If CLng(Me.TextBoxFORTMRD) > 0 Then
            tmp = CLng(Me.TextBoxFORTMRD)
            tmp = tmp - 1
            Me.TextBoxFORTMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForTMRDMore_Click()
    If IsNumeric(Me.TextBoxFORTMRD) Then
        tmp = CLng(Me.TextBoxFORTMRD)
        tmp = tmp + 1
        Me.TextBoxFORTMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForTSMRDLess_Click()
    If IsNumeric(Me.TextBoxForTSMRD) Then
        If CLng(Me.TextBoxForTSMRD) > 0 Then
            tmp = CLng(Me.TextBoxForTSMRD)
            tmp = tmp - 1
            Me.TextBoxForTSMRD = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub ForTSMRDMore_Click()
    If IsNumeric(Me.TextBoxForTSMRD) Then
        tmp = CLng(Me.TextBoxForTSMRD)
        tmp = tmp + 1
        Me.TextBoxForTSMRD = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub HOLess_Click()
    If IsNumeric(Me.TextBoxHO) Then
        If CLng(Me.TextBoxHO) > 0 Then
            tmp = CLng(Me.TextBoxHO)
            tmp = tmp - 1
            Me.TextBoxHO = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub HOMore_Click()
    If IsNumeric(Me.TextBoxHO) Then
        tmp = CLng(Me.TextBoxHO)
        tmp = tmp + 1
        Me.TextBoxHO = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub NALess_Click()
    If IsNumeric(Me.TextBoxNA) Then
        If CLng(Me.TextBoxNA) > 0 Then
            tmp = CLng(Me.TextBoxNA)
            tmp = tmp - 1
            Me.TextBoxNA = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub NAMore_Click()
    If IsNumeric(Me.TextBoxNA) Then
        tmp = CLng(Me.TextBoxNA)
        tmp = tmp + 1
        Me.TextBoxNA = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub OnStockLess_Click()
    If IsNumeric(Me.TextBoxOnStock) Then
        If CLng(Me.TextBoxOnStock) > 0 Then
            tmp = CLng(Me.TextBoxOnStock)
            tmp = tmp - 1
            Me.TextBoxOnStock = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub OnStockMore_Click()
    If IsNumeric(Me.TextBoxOnStock) Then
        tmp = CLng(Me.TextBoxOnStock)
        tmp = tmp + 1
        Me.TextBoxOnStock = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub OpenLess_Click()
    If IsNumeric(Me.TextBoxOPEN) Then
        If CLng(Me.TextBoxOPEN) > 0 Then
            tmp = CLng(Me.TextBoxOPEN)
            tmp = tmp - 1
            Me.TextBoxOPEN = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub OpenMore_Click()
    If IsNumeric(Me.TextBoxOPEN) Then
        tmp = CLng(Me.TextBoxOPEN)
        tmp = tmp + 1
        Me.TextBoxOPEN = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub

Private Sub PotITDCLess_Click()
    If IsNumeric(Me.TextBoxPotITDC) Then
        If CLng(Me.TextBoxPotITDC) > 0 Then
            tmp = CLng(Me.TextBoxPotITDC)
            tmp = tmp - 1
            Me.TextBoxPotITDC = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub PotITDCMore_Click()
    If IsNumeric(Me.TextBoxPotITDC) Then
        tmp = CLng(Me.TextBoxPotITDC)
        tmp = tmp + 1
        Me.TextBoxPotITDC = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub



Private Sub TooLateLess_Click()
    If IsNumeric(Me.TextBoxTooLate) Then
        If CLng(Me.TextBoxTooLate) > 0 Then
            tmp = CLng(Me.TextBoxTooLate)
            tmp = tmp - 1
            Me.TextBoxTooLate = CStr(tmp)
        End If
    End If
    
    zmiany_na_totalach
End Sub

Private Sub TooLateMore_Click()
    If IsNumeric(Me.TextBoxTooLate) Then
        tmp = CLng(Me.TextBoxTooLate)
        tmp = tmp + 1
        Me.TextBoxTooLate = CStr(tmp)
    End If
    
    zmiany_na_totalach
End Sub



Private Sub zmiany_na_totalach()

    ' ale jest to na tyle fajne ze bede mogl dokleic w przyszlosci
    ' inne uby ktore musza sie przeliczac na biezaco
    
    Me.TextBoxTotalFromBoxes.Value = suma_wszystkich_boxow()
    Me.TextBoxTotalFromH1 = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("h1").Value)
    
    
    Dim diff As Long
    diff = 0
    On Error Resume Next
    diff = CLng(Me.TextBoxTotalFromH1) - CLng(Me.TextBoxTotalFromBoxes.Value)
    
    If CLng(diff) > 0 Then
        
        ' Me.TextBoxOpen.Value = diff
        
        Me.TextBoxTotalFromBoxes.BackColor = RGB(255, 255, 0)
        
        ' zmiany_na_totalach
    ElseIf CLng(diff) = 0 Then
        
        Me.TextBoxTotalFromBoxes.BackColor = RGB(0, 255, 0)
    Else
        Me.TextBoxTotalFromBoxes.BackColor = RGB(255, 0, 0)
    End If
End Sub


Private Function suma_wszystkich_boxow() As Long
    suma_wszystkich_boxow = 0
    
    
    sprawdz_czy_pola_przypadkiem_nie_sa_puste
    
    
    suma_wszystkich_boxow = CLng(Me.TextBoxAfterALTMRD.Value) + _
        CLng(Me.TextBoxAfterMRD) + CLng(Me.TextBoxAfterSMRD) + _
        CLng(Me.TextBoxAfterTMRD) + CLng(Me.TextBoxAfterTSMRD) + _
        CLng(Me.TextBoxEDI) + CLng(Me.TextBoxForALTMRD) + _
        CLng(Me.TextBoxForMRD) + CLng(Me.TextBoxFORSMRD) + _
        CLng(Me.TextBoxFORTMRD) + CLng(Me.TextBoxForTSMRD) + _
        CLng(Me.TextBoxHO) + CLng(Me.TextBoxNA) + CLng(Me.TextBoxOnStock) + _
        CLng(Me.TextBoxOPEN) + CLng(Me.TextBoxPotITDC) + CLng(Me.TextBoxTooLate) + _
        CLng(Me.TextBoxForALTTWOMRD) + CLng(Me.TextBoxAfterALTTWOMRD) + _
        CLng(Me.TextBoxForSALTTWOMRD) + CLng(Me.TextBoxAfterSALTTWOMRD) + _
        CLng(Me.TextBoxForONCOSTMRD) + CLng(Me.TextBoxAfterONCOSTMRD) + _
        CLng(Me.TextBoxForSONCOSTMRD) + CLng(Me.TextBoxAfterSONCOSTMRD)
        
        ' new 4 (8) added
        'CLng(Me.TextBoxForALTTWOMRD) + CLng(Me.TextBoxAfterALTTWOMRD) + _
        'CLng(Me.TextBoxForSALTTWOMRD) + CLng(Me.TextBoxAfterSALTTWOMRD) + _
        'CLng(Me.TextBoxForONCOSTMRD) + CLng(Me.TextBoxAfterONCOSTMRD) + _
        'CLng(Me.TextBoxForSONCOSTMRD) + CLng(Me.TextBoxAfterSONCOSTMRD)
        
        
    Me.TextBoxTotalFromBoxes = CLng(suma_wszystkich_boxow)
        
        
End Function

Private Sub sprawdz_czy_pola_przypadkiem_nie_sa_puste()


    Dim item As Control, tb As Control

    For Each item In Me.Controls
    
        If TypeName(item) = "TextBox" Then
        
            
            
            Set tb = item
            
            ' Debug.Print tb.Name
            If tb.Enabled = True Then
                If tb.Value = "" Then
                    tb.Value = 0
                End If
            End If
        End If
    Next item
End Sub

Private Sub UserForm_Activate()
    zmiany_na_totalach
End Sub

Private Sub UserForm_Click()
    zmiany_na_totalach
End Sub

Private Sub UserForm_Initialize()
    zmiany_na_totalach
End Sub
