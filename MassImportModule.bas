Attribute VB_Name = "MassImportModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
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

Public Sub doMassImport(ByRef r As Range)

    ' ------------------------------------------------------------------------------
    ' ------------------------------------------------------------------------------
    ' from new proj form
    ' \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    'r.Value = Me.TextBoxProj
    'r.Offset(0, 1).Value = Me.TextBoxPlt
    'r.Offset(0, 2).Value = Me.TextBoxFaza
    'r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
    'r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
    ' /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
    
    
    ' ORDER RELEASE STATUS
    ' =====================
    orderReleaseStatusMassImportFromWizardBuffer r
    recentBuildPlanChangesMassImportFromWizardBuffer r
    chartContractedPnocMassImportFromWizardBuffer r
    ' osea
    totalsChartMassImportFromWizardBuffer r
    respMassImportFromWizardBuffer r
    delConfStatusMassImportFromWizardBuffer r
    
    
    ' ------------------------------------------------------------------------------
    ' ------------------------------------------------------------------------------
    
End Sub


Private Sub orderReleaseStatusMassImportFromWizardBuffer(ByRef r As Range)
    
    ' sub odpowiadajacy za sciaganie danych z wizard buff worksheet
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    '3: MRD
    '4: BUILD START
    '5: BUILD END
    '6: BOM
    '7: PPAP GATE
    
    With buff
        
        bomFreeze = Replace(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Cells(1, 6)), "CW", ""), "Y", "")
        tmpBuild = Replace(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Cells(1, 4)), "CW", ""), "Y", "")
        tmpMrd = Replace(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Cells(1, 3)), "CW", ""), "Y", "")
        ordersDue = ""
        RELEASED = ""
        numOfVeh = 0
        wksDelay = 0
    End With
    
    Dim orsSh As Worksheet
    Set orsSh = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
    
    
    ' this portion of code similar to inner_calc inside FormOrderReleaseStatus
    ' but only ADD logic
    ' ---------------------------------------------------------------------------
    
    ' no to szukamy pierwszego wolnego pola i wsadzamy
    Dim orsRng As Range
    Set orsRng = orsSh.Cells(1, 1)
    Do
        Set orsRng = orsRng.Offset(1, 0)
    Loop Until Trim(orsRng) = ""
        
    For X = 0 To 3
        orsRng.Offset(0, X) = Trim(r.Offset(0, X))
    Next X
        
        
    ' logic need to be similar to give data to ranges
    ' --------------------------------------------------------------------------------------------
    With orsRng
        .Parent.Cells(.Row, SIXP.e_order_release_mrd) = CStr(tmpMrd)
        .Parent.Cells(.Row, SIXP.e_order_release_build) = CStr(tmpBuild)
        .Parent.Cells(.Row, SIXP.e_order_release_bom_freeze) = CStr(bomFreeze)
        .Parent.Cells(.Row, SIXP.e_order_release_no_of_veh) = CStr(numOfVeh)
        .Parent.Cells(.Row, SIXP.e_order_release_orders_due) = CStr(ordersDue)
        .Parent.Cells(.Row, SIXP.e_order_release_released) = CStr(RELEASED)
        .Parent.Cells(.Row, SIXP.e_order_release_weeks_delay) = CStr(wksDelay)
    End With
    ' --------------------------------------------------------------------------------------------
    
    
    ' after assigning data to order release status sheet we need to add flag on main worksheet
    ' --------------------------------------------------------------------------------------------
    r.Offset(0, SIXP.e_main_last_update_on_order_release_status - 1) = Trim(CStr(r.Offset(0, 3)))
    ' --------------------------------------------------------------------------------------------
        
    
    ' ---------------------------------------------------------------------------
    
End Sub

Private Sub recentBuildPlanChangesMassImportFromWizardBuffer(ByRef r As Range)

    ' ta sekcja pozostanie pusta, poniewaz jako tako nie ma danych w wizard buffer dla niej
    ' ale zeby konsystencja kodu zostala w miare przejrzysta zostawie ten oto szablonik

End Sub

Private Sub chartContractedPnocMassImportFromWizardBuffer(ByRef r As Range)

    
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    Dim rngv As Range, rngl As Range, total_total As Range
    Set rngv = buff.Cells(3, 1)
    Set rngl = buff.Cells(2, 1)
    Set total_total = buff.Range("B4")
    
    
    With buff

        PNOC = CStr(SIXP.GetDataFromWizardBufferModule.get_all_values("PNOC", rngl))
        ' total total
        tot_tot = CStr(CLng(total_total))
        Contracted = CStr(CLng(tot_tot) - CLng(PNOC))
    End With
    
    
    
    ' data to ranges
    ' no to szukamy pierwszego wolnego pola i wsadzamy
    ' ===================================================
    Dim ccp As Range
    Set ccp = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm).Cells(1, 1)
    Do
        Set ccp = ccp.Offset(1, 0)
    Loop Until Trim(ccp) = ""
    
    For X = 0 To 3
        ccp.Offset(0, X) = Trim(r.Offset(0, X))
    Next X
    
    
    ' assign to ranges
    ' --------------------------------------------------------------------------------------------
    With ccp
        .Parent.Cells(.Row, SIXP.e_cont_pnoc_chart_actionable_fma) = CStr(0)
        .Parent.Cells(.Row, SIXP.e_cont_pnoc_chart_contracted) = CStr(Contracted)
        .Parent.Cells(.Row, SIXP.e_cont_pnoc_chart_open_bp) = CStr(0)
        .Parent.Cells(.Row, SIXP.e_cont_pnoc_chart_pnoc) = CStr(PNOC)
    End With
    ' --------------------------------------------------------------------------------------------
    
    
    ' update main sh
    ' --------------------------------------------------------------------------------------------
    r.Offset(0, SIXP.e_main_last_update_on_chart_contracted_pnoc - 1) = Trim(CStr(r.Offset(0, 3)))
    ' --------------------------------------------------------------------------------------------
    
    ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
    ' dodanie info na sam koniec tabeli
    
    
    
    ' ===================================================
    
End Sub



Private Sub totalsChartMassImportFromWizardBuffer(ByRef r As Range)


    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    '3: MRD
    '4: BUILD START
    '5: BUILD END
    '6: BOM
    '7: PPAP GATE
    
    total_num = 0
    If IsNumeric(buff.Cells(1, 8)) Then
        total_num = CLng(buff.Cells(1, 8))
    End If
    
    With buff
        
        
        ' porcja zwiazana z totalami eur i osea
        ' ---------------------------------------------------------------
        osea_v = 0
        eur_v = 0
        
        If IsNumeric(.Cells(32, 1)) Then osea_v = .Cells(32, 1)
        If IsNumeric(.Cells(32, 2)) Then eur_v = .Cells(32, 2)
        ' ---------------------------------------------------------------
        
        
        ' porcja total z perspektywy transportow
        ' ---------------------------------------------------------------
        ARRIVED = 0
        in_t = 0
        future = 0
        
        If IsNumeric(.Cells(37, 1)) Then ARRIVED = .Cells(37, 1)
        If IsNumeric(.Cells(37, 2)) Then in_t = .Cells(37, 2)
        If IsNumeric(.Cells(37, 3)) Then future = .Cells(37, 3)
        ' ---------------------------------------------------------------
        
        ' porcja PNOC
        ' nieco bardziej zlozona bo trzeba jeszcze znalezc slowo
        ' klucz w wierszy drugim, liczba w wierszu 3
        ' ---------------------------------------------------------------
        pnoc_v = 0
        pnoc_v = ktora_kolumna_to__foo(.Cells(2, 1), "PNOC")
        ' ---------------------------------------------------------------
        
        
        ' ---------------------------------------------------------------
        itdc_v = 0
        itdc_v = ktora_kolumna_to__foo(.Cells(2, 1), "ITDC")
        ' ---------------------------------------------------------------
        
        ' ---------------------------------------------------------------
        mpc_v = 0
        mpc_v = ktora_kolumna_to__foo(.Cells(2, 1), "MPC")
        ' ---------------------------------------------------------------
        
        ' ---------------------------------------------------------------
        ordered_v = 0
        ordered_v = ktora_kolumna_to__foo(.Cells(41, 1), "OK")
        ' ---------------------------------------------------------------
        
        
        ' PPAP
        ' na ppap skladaja sie ok i nok i koncepcja obliczeniowa jest nieco bardziej skomplikowana
        ' bedzie trzeba sie posilkowac lista w rejestrze, ktore stringi sa ok, a ktore sa nok
        ' ---------------------------------------------------------------
        
        ppap_ok = 0
        ppap_nok = 0
        
        ppap_ok = get_ppaps_foo(.Cells(6, 1), E_PPAP_OK)
        ppap_nok = get_ppaps_foo(.Cells(6, 1), E_PPAP_NOK)
        
        ' ---------------------------------------------------------------
    End With
    
    
    ' teraz wyliczymy NA na podstawie roznicy totala z juz zaciagnietych danych
    na_v = 0
    '!
    
    na_v = CLng(total_num - pnoc_v - itdc_v - osea_v - eur_v)
    
    If CLng(na_v) < 0 Then
        na_v = 0
    End If
    
    

    
    ' data to ranges
    ' no to szukamy pierwszego wolnego pola i wsadzamy
    ' ===================================================
    Dim tot As Range
    Set tot = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
    Do
        Set tot = tot.Offset(1, 0)
    Loop Until Trim(tot) = ""
    
    For X = 0 To 3
        tot.Offset(0, X) = Trim(r.Offset(0, X))
    Next X
    
    
    ' assign to ranges
    ' --------------------------------------------------------------------------------------------
    With tot
        
        .Parent.Cells(.Row, SIXP.e_5p_total) = CStr(total_num)
        
        .Parent.Cells(.Row, SIXP.e_5p_fma_eur) = CStr(eur_v)
        .Parent.Cells(.Row, SIXP.e_5p_fma_osea) = CStr(osea_v)
        
        .Parent.Cells(.Row, SIXP.e_5p_arrived) = CStr(ARRIVED)
        .Parent.Cells(.Row, SIXP.e_5p_future) = CStr(future)
        .Parent.Cells(.Row, SIXP.e_5p_in_transit) = CStr(in_t)
        
        .Parent.Cells(.Row, SIXP.e_5p_itdc) = CStr(itdc_v)
        .Parent.Cells(.Row, SIXP.e_5p_na) = CStr(na_v)
        
        .Parent.Cells(.Row, SIXP.e_5p_no_ppap_status) = CStr(ppap_nok)
        .Parent.Cells(.Row, SIXP.e_5p_ppap_status) = CStr(ppap_ok)
        
        .Parent.Cells(.Row, SIXP.e_5p_arrived) = CStr(ARRIVED)
        .Parent.Cells(.Row, SIXP.e_5p_arrived) = CStr(ARRIVED)
        .Parent.Cells(.Row, SIXP.e_5p_arrived) = CStr(ARRIVED)
        .Parent.Cells(.Row, SIXP.e_5p_ordered) = CStr(ordered_v)
        .Parent.Cells(.Row, SIXP.e_5p_pnoc) = CStr(pnoc_v)
    End With
    ' --------------------------------------------------------------------------------------------
    
    
    ' update main sh
    ' --------------------------------------------------------------------------------------------
    r.Offset(0, SIXP.e_main_last_update_on_totals - 1) = Trim(CStr(r.Offset(0, 3)))
    ' --------------------------------------------------------------------------------------------

End Sub



' duplikat funkcji wystepujacy rowniez w formularzu total
' jednak zasieg ten jest mocno odseparowany wiec narazie
' zrobie tak ze ten sam kod zamieszcze dwa razy :(
Private Function ktora_kolumna_to__foo(r As Range, txt As String) As Long
    
    ktora_kolumna_to__foo = 0
    
    Do
        If Trim(r) = CStr(txt) Then
            ktora_kolumna_to__foo = CLng(r.Offset(1, 0))
            Exit Function
        End If
        
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
End Function


' duplikat funkcji wystepujacy rowniez w formularzu total
' jednak zasieg ten jest mocno odseparowany wiec narazie
' zrobie tak ze ten sam kod zamieszcze dwa razy :(
Private Function get_ppaps_foo(r As Range, e As E_PPAP) As Long
    get_ppaps_foo = 0
    
    
    ' 2 petle
    ' jedna po wiz buff
    ' druga po register list
    
    ' range wizz buff
    Dim rwb As Range
    ' range register
    Dim rr As Range, tmp_r As Range
    
    Set rwb = r
    
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("J2")
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range(rr, rr.End(xlDown))
    
    Do
        For Each tmp_r In rr
            If Trim(tmp_r) = Trim(rwb) Then
                If CLng(tmp_r.Offset(0, 1)) = CLng(e) Then
                    
                    ' ----------------------------------------------------------------
                    ' ----------------------------------------------------------------
                    ''
                    '
                    get_ppaps_foo = get_ppaps_foo + CLng(rwb.Offset(1, 0))
                    '
                    ''
                    ' ----------------------------------------------------------------
                    ' ----------------------------------------------------------------
                End If
            End If
        Next tmp_r
        Set rwb = rwb.Offset(0, 1)
    Loop Until Trim(rwb) = ""
    
End Function


Private Sub delConfStatusMassImportFromWizardBuffer(ByRef r As Range)



    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    
    Dim rngv As Range, rngl As Range, h1_total
    Set rngv = buff.Cells(17, 1)
    Set rngl = buff.Cells(16, 1)
    Set h1_total = buff.Range("H1")
    
    
    With buff
    
        ' 15: BEFORE OR ON/AFTER MRD
        ' 16: BEFORE/AFTER MRD - labels all
        ' 17: values
        afterALTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_MRD))
        
        forALTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_MRD))
        
        afterMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_MRD))
        forMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_MRD))
        
        afterSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_MRD))
        forSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_MRD))
        
        'afterTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_TWO_MRD))
        'forTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_TWO_MRD))
        
        'afterTSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_TWO_Staggered_MRD))
        'forTSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_TWO_Staggered_MRD))
        
        
        
        ' new
        ' ---------------------------------------------------------------------------------------------------------------------
        
        afterALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_TWO_MRD))
        forALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_TWO_MRD))
        
        afterSALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        forSALTTWOMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        
        afterONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ONCOST_MRD))
        forONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ONCOST_MRD))
        
        afterSONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_ONCOST_MRD))
        forSONCOSTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_ONCOST_MRD))
        
        ' ---------------------------------------------------------------------------------------------------------------------
    End With
    
    afterALTMRD = emptyToZero(afterALTMRD)
    forALTMRD = emptyToZero(forALTMRD)
    afterMRD = emptyToZero(afterMRD)
    forMRD = emptyToZero(forMRD)
    afterSMRD = emptyToZero(afterSMRD)
    forSMRD = emptyToZero(forSMRD)
    afterALTTWOMRD = emptyToZero(afterALTTWOMRD)
    forALTTWOMRD = emptyToZero(forALTTWOMRD)
    afterSALTTWOMRD = emptyToZero(afterSALTTWOMRD)
    forSALTTWOMRD = emptyToZero(forrSALTTWOMRD)
    afterONCOSTMRD = emptyToZero(afterONCOSTMRD)
    forONCOSTMRD = emptyToZero(forONCOSTMRD)
    afterSONCOSTMRD = emptyToZero(afterSONCOSTMRD)
    forSONCOSTMRD = emptyToZero(forSONCOSTMRD)

    
    ' DEL CONF, WHICH IS NOT CONNECTED WITH MRD PARAM.
    
    Set rngv = buff.Cells(13, 1)
    Set rngl = buff.Cells(12, 1)
    
    With buff
        
        
        ' greens
        onStock = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_ON_STOCK))
        edi = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_EDI))
        ho = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_HO))
        NA = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_NA))
        
        onStock = emptyToZero(onStock)
        edi = emptyToZero(edi)
        ho = emptyToZero(ho)
        
        
        
        'reds
        ' jednak tutaj jest powazny problem poniewaz wizard jako tako nie bierze pod uwagi podzialu NOKow
        ' traktowane sa one normlanie jako blanki bez wiekszego zglebiania
        ' zatem ponizsza logika nie ma sensu zeby byla powielana w wykorzystaniu tak jak to mialo miejsce w greensach
        ' czy polach uzaleznionych od MRD
        ' me.TextBoxOpen = cstr(sixp.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl...)
        openStr = 0
        tooLateStr = 0
        ' wyjatekiem jest pot itdc, poniewaz jako tako mozna wyrazic ten element za pomoca stringa zamieszcznego w wizardzie
        potItdc = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_POTITDC))
        potItdc = emptyToZero(potItdc)
    End With
    
    
    Dim suma_wszystkich_boxow As Long
    suma_wszystkich_boxow = 0 + afterALTMRD + forALTMRD + afterMRD + forMRD + afterSMRD + forSMRD + _
        afterALTTWOMRD + forALTTWOMRD + afterSALTTWOMRD + forSALTTWOMRD + afterONCOSTMRD + forONCOSTMRD + _
        afterSONCOSTMRD + forSONCOSTMRD + potItdc + tooLateStr

        
    openStr = CStr(CLng(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("h1").Value) - suma_wszystkich_boxow)




    ' ===================================================
    Dim dcRng As Range
    Set dcRng = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
    Do
        Set dcRng = dcRng.Offset(1, 0)
    Loop Until Trim(dcRng) = ""
    
    

    
    For X = 0 To 3
        dcRng.Offset(0, X) = Trim(r.Offset(0, X))
    Next X
    
    
    ' give_data_to_ranges r - adapt
    ' --------------------------------------------------------------------------------------------------
    
    With dcRng
        .Parent.Cells(.Row, SIXP.e_del_conf_on_stock) = CStr(onStock)
        .Parent.Cells(.Row, SIXP.e_del_conf_edi) = CStr(edi)
        .Parent.Cells(.Row, SIXP.e_del_conf_edi) = CStr(ho)
        .Parent.Cells(.Row, SIXP.e_del_conf_na) = CStr(NA)
        
        ' mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_mrd) = CStr(forMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_mrd) = CStr(afterMRD)
        
        ' staggered mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_smrd) = CStr(forSMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_smrd) = CStr(afterSMRD)
        
        ' alt mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_alt) = CStr(forALTMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_alt) = CStr(afterALTMRD)
        
        ' alt two mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_alttwomrd) = CStr(forALTTWOMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_alttwomrd) = CStr(afterALTTWOMRD)
        
        ' staggered alt two mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_salttwomrd) = CStr(forSALTTWOMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_salttwomrd) = CStr(afterSALTTWOMRD)
        
        ' on cost mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_oncostmrd) = CStr(forONCOSTMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_oncostmrd) = CStr(afterONCOSTMRD)
        
        ' staggered on cost mrd
        .Parent.Cells(.Row, SIXP.e_del_conf_for_soncostmrd) = CStr(forSONCOSTMRD)
        .Parent.Cells(.Row, SIXP.e_del_conf_after_soncostmrd) = CStr(afterSONCOSTMRD)
        
        ' open too late pot itdc
        .Parent.Cells(.Row, SIXP.e_del_conf_open) = CStr(openStr)
        .Parent.Cells(.Row, SIXP.e_del_conf_too_late) = CStr(tooLateStr)
        .Parent.Cells(.Row, SIXP.e_del_conf_pot_itdc) = CStr(potItdc)
    End With
    

    ' --------------------------------------------------------------------------------------------------
    
    
    
    ' change_col_L_in_MAIN_worksheet r
    ' jest to samo w order release status sheet oraz to samo w main sheet
    ' --------------------------------------------------------------------
    ''
    '
    r.Offset(0, SIXP.e_main_last_update_on_del_conf - 1) = Trim(CStr(r.Offset(0, 3)))
    '
    ''
    ' --------------------------------------------------------------------

    
    
    
    ' ===================================================
End Sub

Private Function emptyToZero(str) As Integer
    
    If Trim(CStr(str)) = "" Then
        emptyToZero = 0
    Else
        If IsNumeric(str) Then
            emptyToZero = CLng(str)
        Else
            emptyToZero = 0
        End If
    End If
End Function




Private Sub respMassImportFromWizardBuffer(ByRef r As Range)
    
    
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    nm = CStr(buff.Range("J1"))
    
    
    ' ===================================================
    Dim respRng As Range
    Set respRng = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm).Cells(1, 1)
    Do
        Set respRng = respRng.Offset(1, 0)
    Loop Until Trim(respRng) = ""
    
    

    
    For X = 0 To 3
        respRng.Offset(0, X) = Trim(r.Offset(0, X))
    Next X
    
    ' ===================================================
    
    
    ' give_data_to_ranges r - adapt
    ' --------------------------------------------------------------------------------------------------
    
    With respRng
    
        .Parent.Cells(.Row, SIXP.e_resp_fma) = CStr(nm)
    End With
    ' --------------------------------------------------------------------------------------------------
    
    
    ' update Main Sheet to last update on resp (N column in this case)
    r.Offset(0, SIXP.e_main_last_update_on_resp - 1) = Trim(CStr(r.Offset(0, 3)))
    
End Sub
