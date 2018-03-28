VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewDelConfHandler"
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



Private m As Worksheet
Private delConf As Worksheet

Private r As Range

Private Sub Class_Initialize()

    Application.EnableEvents = False
    
    ' sheets assigning
    ' ====================================================================
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    Set delConf = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set delConf = Nothing
    
    Application.EnableEvents = True
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormDelConfStatus2)


    Application.EnableEvents = False
    
    
    GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = False
    
    If CStr(eca) = SIXP.G_BTN_TEXT_ADD Then
    
        With o
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            ' zerowanie znajduje sie w innej czesci implementacji...
            
            'Dim c As Control
            'For Each c In .Controls
            '
            '    If c.name Like "TextBox*" Then
            '        c.Value = 0
            '    End If
            'Next c
            
            
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


                        '.TextBoxAfterALTMRD = _
                        '    SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_alt - 1)))
                        ' new
                        '.TextBoxAfterALTTWOMRD = _
                        '    SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_alttwomrd - 1)))
                        
                        
                        ' green szajs
                        ' -----------------------------------------------------------------------------------------------
                        .TextBoxGALT = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_alt - 1)))
                        .TextBoxGALTTWO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_alttwomrd - 1)))
                        .TextBoxGEDI = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_edi - 1)))
                        .TextBoxGHO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_ho - 1)))
                        .TextBoxGMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_mrd - 1)))
                        .TextBoxGNA = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_na - 1)))
                        .TextBoxGONCOST = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_oncostmrd - 1)))
                        .TextBoxGONSTOCK = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_on_stock - 1)))
                        .TextBoxGSALTTWO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_salttwomrd - 1)))
                        .TextBoxGSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_smrd - 1)))
                        .TextBoxGSONCOST = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_for_soncostmrd - 1)))
                        ' -----------------------------------------------------------------------------------------------
                        
                        
                        ' yellow szajs
                        ' -----------------------------------------------------------------------------------------------
                        .TextBoxYALT = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_alt - 1)))
                        .TextBoxYALTTWO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_alttwomrd - 1)))
                        .TextBoxYMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_mrd - 1)))
                        .TextBoxYONCOST = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_oncostmrd - 1)))
                        .TextBoxYOpen = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_yellow_open - 1)))
                        .TextBoxYSALTTWO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_salttwomrd - 1)))
                        .TextBoxYSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_smrd - 1)))
                        .TextBoxYSONCOST = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_after_soncostmrd - 1)))
                        
                        
                        ' -----------------------------------------------------------------------------------------------
                        
                        
                        ' red szajs
                        ' -----------------------------------------------------------------------------------------------
                        
                        .TextBoxRALT = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_alt - 1)))
                        .TextBoxRALTTWO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_alttwomrd - 1)))
                        .TextBoxRMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_mrd - 1)))
                        .TextBoxRONCOST = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_oncostmrd - 1)))
                        .TextBoxROpen = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_open - 1)))
                        .TextBoxRPOTITDC = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_pot_itdc - 1)))
                        .TextBoxRSALTTWO = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_salttwomrd - 1)))
                        .TextBoxRSMRD = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_smrd - 1)))
                        .TextBoxRSONCOST = _
                            SIXP.GlobalFooModule.global_cpz(CStr(r.Offset(0, SIXP.e_del_conf_red_after_soncostmrd - 1)))
                        
                        
                        ' -----------------------------------------------------------------------------------------------
                        
                        
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
        End With
    End If
    
    Application.EnableEvents = True
    
    
    GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = True
End Sub
