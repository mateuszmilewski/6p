VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DelConfStatus7XHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta ma sie zajmowac kontrola danych miedzy order releasami a main ogolnie
' dobrze by bylo potrfila wpisywac ladnie aktualny ycw dla update'u arkusza
' order release, dzieki czemu bedziemy mogli miec porownanie z aktualnym cw w arkuszu main

Private m As Worksheet
Private t As Worksheet

Private r As Range

Private Sub Class_Initialize()


    ' ten pseudo konstruktor nie bedzie mial narazie wiecej kodu
    ' 2016-08-10

    ' main sheets assigning
    ' ====================================================================
    ' przypisanie arkusza main
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' arkusz order release status handler
    Set t = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set t = Nothing
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormDelConfStatus)


    If CStr(eca) = SIXP.G_BTN_TEXT_ADD Then
    
        With o
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            
            'Public Enum E_RECENT_BP_CH
            '    e_recent_bp_ch_no_of_veh = 5
            '    e_recent_bp_ch_tbw
            '    e_recent_bp_ch_order_release_ch
            '    e_recent_bp_ch_comment
            'End Enum
            
            
            ' pola liczbowe
            ' ---------------------------------------------------------
            .TextBoxAfterALTMRD = "0"
            .TextBoxAfterMRD = "0"
            .TextBoxAfterSMRD = "0"
            .TextBoxAfterTMRD = "0"
            .TextBoxAfterTSMRD = "0"
            .TextBoxEDI = "0"
            .TextBoxForALTMRD = "0"
            .TextBoxForMRD = "0"
            .TextBoxForSMRD = "0"
            .TextBoxForTMRD = "0"
            .TextBoxForTSMRD = "0"
            .TextBoxHO = "0"
            .TextBoxNA = "0"
            .TextBoxOnStock = "0"
            .TextBoxOPEN = "0"
            .TextBoxPotITDC = "0"
            .TextBoxUNDEF = "0"
            
            ' ---------------------------------------------------------
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            Set r = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
            Do
                If CStr(.LabelTitle.Caption) = _
                    CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                

                        
                        ' pola liczbowe
                        ' ---------------------------------------------------------


                        .TextBoxAfterALTMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_alt - 1))
                        .TextBoxAfterMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_mrd - 1))
                        .TextBoxAfterSMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_smrd - 1))
                        .TextBoxAfterTMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_twomrd - 1))
                        .TextBoxAfterTSMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_twosmrd - 1))
                        .TextBoxEDI = CStr(r.Offset(0, SIXP.e_del_conf_edi - 1))
                        .TextBoxForALTMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_alt - 1))
                        .TextBoxForMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_mrd - 1))
                        .TextBoxForSMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_smrd - 1))
                        .TextBoxForTMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_twomrd - 1))
                        .TextBoxForTSMRD = CStr(r.Offset(0, SIXP.e_del_conf_for_twosmrd - 1))
                        .TextBoxHO = CStr(r.Offset(0, SIXP.e_del_conf_ho - 1))
                        .TextBoxNA = CStr(r.Offset(0, SIXP.e_del_conf_na - 1))
                        .TextBoxOnStock = CStr(r.Offset(0, SIXP.e_del_conf_on_stock - 1))
                        .TextBoxOPEN = CStr(r.Offset(0, SIXP.e_del_conf_open - 1))
                        .TextBoxPotITDC = CStr(r.Offset(0, SIXP.e_del_conf_pot_itdc - 1))
                        .TextBoxUNDEF = CStr(r.Offset(0, SIXP.e_del_conf_undef - 1))
                        
                        ' ---------------------------------------------------------
                    
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
        End With
    End If
End Sub






