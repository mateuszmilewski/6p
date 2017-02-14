VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DelConfStatus7XHandler"
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
            .TextBoxOpen = "0"
            .TextBoxPotITDC = "0"
            .TextBoxTooLate = "0"
            
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


                        .TextBoxAfterALTMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_alt - 1)))
                        .TextBoxAfterMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_mrd - 1)))
                        .TextBoxAfterSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_smrd - 1)))
                        .TextBoxAfterTMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_twomrd - 1)))
                        .TextBoxAfterTSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_twosmrd - 1)))
                        .TextBoxEDI = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_edi - 1)))
                        .TextBoxForALTMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_alt - 1)))
                        .TextBoxForMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_mrd - 1)))
                        .TextBoxForSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_smrd - 1)))
                        .TextBoxForTMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_twomrd - 1)))
                        .TextBoxForTSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_twosmrd - 1)))
                        .TextBoxHO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_ho - 1)))
                        .TextBoxNA = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_na - 1)))
                        .TextBoxOnStock = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_on_stock - 1)))
                        .TextBoxOpen = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_open - 1)))
                        .TextBoxPotITDC = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_pot_itdc - 1)))
                        .TextBoxTooLate = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_too_late - 1)))
                        
                        ' ---------------------------------------------------------
                    
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
        End With
    End If
End Sub






