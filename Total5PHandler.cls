VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Total5PHandler"
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
    Set t = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set t = Nothing
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormTotals5p)


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
            .TextBoxArrived = "0"
            .TextBoxFMAEUR = "0"
            .TextBoxFmaOsea = "0"
            .TextBoxInTransit = "0"
            .TextBoxITDC = "0"
            .TextBoxNA = "0"
            .TextBoxNoPPAP = "0"
            .TextBoxOrdered = "0"
            .TextBoxPNOC = "0"
            .TextBoxPPAP = "0"
            .TextBoxTotal = "0"
            
            ' ---------------------------------------------------------
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            Set r = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
            Do
                If CStr(.LabelTitle.Caption) = _
                    CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                

                        
                        ' pola liczbowe
                        ' ---------------------------------------------------------

                        .TextBoxArrived = CStr(r.Offset(0, SIXP.e_5p_arrived - 1))
                        .TextBoxFMAEUR = CStr(r.Offset(0, SIXP.e_5p_fma_eur - 1))
                        .TextBoxFmaOsea = CStr(r.Offset(0, SIXP.e_5p_fma_osea - 1))
                        .TextBoxInTransit = CStr(r.Offset(0, SIXP.e_5p_in_transit - 1))
                        .TextBoxITDC = CStr(r.Offset(0, SIXP.e_5p_itdc - 1))
                        .TextBoxInTransit = CStr(r.Offset(0, SIXP.e_5p_na - 1))
                        .TextBoxNoPPAP = CStr(r.Offset(0, SIXP.e_5p_no_ppap_status - 1))
                        .TextBoxOrdered = CStr(r.Offset(0, SIXP.e_5p_ordered - 1))
                        .TextBoxPNOC = CStr(r.Offset(0, SIXP.e_5p_pnoc - 1))
                        .TextBoxPPAP = CStr(r.Offset(0, SIXP.e_5p_ppap_status - 1))
                        .TextBoxTotal = CStr(r.Offset(0, SIXP.e_5p_total - 1))
                        
                        ' ---------------------------------------------------------
                    
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            
            
        
        End With
    End If
End Sub




