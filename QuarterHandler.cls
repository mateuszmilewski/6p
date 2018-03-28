VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuarterHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private d As Dictionary
Private tl As T_Link
Private l As Linker
Private r As Range
Private Sh As Worksheet

Private nazwaPliku As String


Public Function getQuarterMainTable() As Worksheet
    Set getQuarterMainTable = Sh
End Function

Public Function zrobKlucz(strProj, strPlt, strFaza, strCW)
    Set tl = New T_Link
    tl.zrob_mnie_z_argsow strProj, strPlt, strFaza, strCW
    zrobKlucz = CStr(r.Row) & "_" & CStr(Replace(CStr(l.return_full_concated_r_string(tl)), "_", ""))
End Function

Public Sub fillDictionaryWithTLinks(ByRef msh As Worksheet)


    Set Sh = msh

    Set r = Sh.Cells(2, SIXP.e_qmt_project)
    
    If Trim(r) <> "" Then
    
    Do
        
        klucz = zrobKlucz(r.Value, Sh.Cells(r.Row, SIXP.e_qmt_Plant).Value, Sh.Cells(r.Row, SIXP.e_qmt_PHASE).Value, Sh.Cells(r.Row, SIXP.e_qmt_cw).Value)
        
        If Not d.Exists(klucz) Then
            d.Add klucz, tl
        Else
            ' nie do konca ta logika agresywna sie sprawdza - ostatecznie zrobimy to tak ze bede nadpisywal dane
            ' MsgBox "masz duplikaty w swoim Quarterze - popraw to!"
            
            Set d(klucz) = tl
        End If
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Else
        MsgBox "Quarter nie ma danych! Apka konczy dzialanie!"
        End
    End If
    

End Sub


Public Sub openFormWithDataFromQuarter()
    
    fillListBoxes
    noSelectionOnListBoxes
    
    
    FormQuarterContentChooser.passData d, Me, nazwaPliku
    FormQuarterContentChooser.Show vbModeless
    
    
    
End Sub

Public Sub noSelectionOnListBoxes()

    With FormQuarterContentChooser
    
        For x = 0 To .ListBoxProj.ListCount - 1
            
            .ListBoxProj.Selected(x) = False
            
        Next x
    End With
End Sub

Public Sub resetListBoxes()
    fillListBoxes
End Sub

Public Sub wildcardListBoxes(wc_proj, wc_plt, wc_faza, wc_cw)

    With FormQuarterContentChooser
        .ListBoxCW.Clear
        .ListBoxFaza.Clear
        .ListBoxPLT.Clear
        .ListBoxProj.Clear
        .ListBoxRow.Clear
    End With


    If d.Count > 0 Then
    
        For Each Key In d.Keys
    
            With FormQuarterContentChooser
            
                Set tl = d(Key)
                
                wiersz = Split(Key, "_")(0)
                
                If UCase(tl.project) Like "*" & UCase(wc_proj) & "*" Then
                    If UCase(tl.plt) Like "*" & UCase(wc_plt) & "*" Then
                        If UCase(tl.faza) Like "*" & UCase(wc_faza) & "*" Then
                            If UCase(tl.cw) Like "*" & UCase(wc_cw) & "*" Then
                                .ListBoxRow.AddItem wiersz
                                .ListBoxProj.AddItem tl.project
                                .ListBoxPLT.AddItem tl.plt
                                .ListBoxFaza.AddItem tl.faza
                                .ListBoxCW.AddItem tl.cw
                            End If
                        End If
                    End If
                End If
            
            End With
        
        Next
    
    Else
        MsgBox "pojawil sie blad braku danych pobranych z Quartera, makro konczy dzialanie w nieoczekiwany sposob!"
        End
    End If
End Sub

Private Sub fillListBoxes()


    With FormQuarterContentChooser
        .ListBoxCW.Clear
        .ListBoxFaza.Clear
        .ListBoxPLT.Clear
        .ListBoxProj.Clear
    End With


    If d.Count > 0 Then
    
        
    
    
        For Each Key In d.Keys
    
            With FormQuarterContentChooser
            
                Set tl = d(Key)
                wiersz = Split(Key, "_")(0)
                .ListBoxRow.AddItem wiersz
                .ListBoxProj.AddItem tl.project
                .ListBoxPLT.AddItem tl.plt
                .ListBoxFaza.AddItem tl.faza
                .ListBoxCW.AddItem tl.cw
                
            
            End With
        
        Next
    
    Else
        MsgBox "pojawil sie blad braku danych pobranych z Quartera, makro konczy dzialanie w nieoczekiwany sposob!"
        End
    End If
End Sub

Public Sub overwriteDictionary(newD As Dictionary)
    Set d = newD
End Sub

Private Sub Class_Initialize()
    
    Set d = New Dictionary
    Set l = New Linker
    Set Sh = Nothing
    
    nazwaPliku = ""
    
    
    SIXP.LoadingFormModule.increaseLoadingFormStatus 50
End Sub

Public Sub setWrkNm(wrknm)

    SIXP.LoadingFormModule.increaseLoadingFormStatus 50
    
    nazwaPliku = CStr(wrknm)
End Sub

Public Function getWrkNm()
    getWrkNm = CStr(nazwaPliku)
End Function

Private Sub Class_Terminate()
    Set d = Nothing
    Set l = Nothing
    Set Sh = Nothing
    
    SIXP.LoadingFormModule.increaseLoadingFormStatus 50
    
    MsgBox "import Quarter: ready!"
    
End Sub



Public Sub copyDataFromQuarter(mt As Worksheet, c As Collection)


    ' pierwsza "gitara" - to znaczy ze ok ze mozemy wpisywac dane do 6P z Quartera
    If c.Count > 0 Then
        
        
        Dim r As Range
        
        For Each w In c
        
            SIXP.LoadingFormModule.incLoadingForm
        
            Set r = mt.Cells(w, SIXP.e_qmt_project)
            Set r = mt.Range(r, mt.Cells(w, SIXP.e_qmt_after_MRD_osea))
            
            ' Debug.Print mt.Cells(w, SIXP.e_qmt_project).Value & " " & mt.Cells(w, SIXP.e_qmt_MRD).Value
            
            
            
            
            SIXP.LoadingFormModule.incLoadingForm
            sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_main_sh_nm))
            ' dodatkowa zmienna, by moc do niej wrocic na koniec logiki, gdy pozostale arkusze zostana juz uzupelnione
            main_w = sixp_w
            With ThisWorkbook
                With .Sheets(SIXP.G_main_sh_nm)
                
                
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt
                    
                    ' status niby nalezy do main jednak nie jest czescia odpowiedzialna za linkowanie danych
                    ' ----------------------------------------------------------------------------------------
                    .Cells(sixp_w, SIXP.e_main_status).Value = mt.Cells(w, SIXP.e_qmt_status).Value
                    ' ----------------------------------------------------------------------------------------
                    
                    ' reszta danych zostanie uzupelniona potem, gdy dane na innych arkuszach sie juz pojawia
                End With
                
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm))
                With .Sheets(SIXP.G_order_release_status_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt
                    
                    'Public Enum E_ORDER_RELEASE_STATUS
                    '    e_order_release_mrd = 5
                    '    e_order_release_build
                    '    e_order_release_bom_freeze
                    '    e_order_release_no_of_veh
                    '    e_order_release_orders_due
                    '    e_order_release_released
                    '    e_order_release_weeks_delay
                    'End Enum
                    
                    
                    
                    .Cells(sixp_w, SIXP.e_order_release_mrd).Value = mt.Cells(w, SIXP.e_qmt_MRD).Value
                    .Cells(sixp_w, SIXP.e_order_release_build).Value = mt.Cells(w, SIXP.e_qmt_Build_ORS).Value
                    .Cells(sixp_w, SIXP.e_order_release_bom_freeze).Value = mt.Cells(w, SIXP.e_qmt_BOM_freeze).Value
                    .Cells(sixp_w, SIXP.e_order_release_no_of_veh).Value = mt.Cells(w, SIXP.e_qmt_NO_OF_VEHICLES_ORS).Value
                    .Cells(sixp_w, SIXP.e_order_release_orders_due).Value = mt.Cells(w, SIXP.e_qmt_ordersDUE).Value
                    .Cells(sixp_w, SIXP.e_order_release_released).Value = mt.Cells(w, SIXP.e_qmt_RELEASED).Value
                    .Cells(sixp_w, SIXP.e_order_release_weeks_delay).Value = mt.Cells(w, SIXP.e_qmt_WeeksDelay).Value

                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm))
                With .Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt
                    

                    'Public Enum E_RECENT_BP_CH
                    '    e_recent_bp_ch_no_of_veh = 5
                    '    e_recent_bp_ch_tbw
                    '    e_recent_bp_ch_order_release_ch
                    '    e_recent_bp_ch_comment
                    'End Enum
                    
                    .Cells(sixp_w, SIXP.e_recent_bp_ch_no_of_veh).Value = mt.Cells(w, SIXP.e_qmt_NO_OF_VEHICLES_RBPC).Value
                    .Cells(sixp_w, SIXP.e_recent_bp_ch_tbw).Value = mt.Cells(w, SIXP.e_qmt_TBW).Value
                    .Cells(sixp_w, SIXP.e_recent_bp_ch_order_release_ch).Value = mt.Cells(w, SIXP.e_qmt_OrderReleaseChange_RBPC).Value
                    .Cells(sixp_w, SIXP.e_recent_bp_ch_comment).Value = mt.Cells(w, SIXP.e_qmt_Comments_RBPC).Value
                    
                End With
                
                SIXP.LoadingFormModule.incLoadingForm
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm))
                With .Sheets(SIXP.G_cont_pnoc_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt
                    

                    'Public Enum E_CONT_PNOC_CHART
                    '    e_cont_pnoc_chart_contracted = 5
                    '    e_cont_pnoc_chart_pnoc
                    '    e_cont_pnoc_chart_open_bp
                    '    e_cont_pnoc_chart_actionable_fma
                    'End Enum
                    
                    .Cells(sixp_w, SIXP.e_cont_pnoc_chart_contracted).Value = mt.Cells(w, SIXP.e_qmt_Contracted).Value
                    .Cells(sixp_w, SIXP.e_cont_pnoc_chart_pnoc).Value = mt.Cells(w, SIXP.e_qmt_PNOC_Chart).Value
                    .Cells(sixp_w, SIXP.e_cont_pnoc_chart_open_bp).Value = mt.Cells(w, SIXP.e_qmt_OpenBP_Chart).Value
                    .Cells(sixp_w, SIXP.e_cont_pnoc_chart_actionable_fma).Value = mt.Cells(w, SIXP.e_qmt_actionableFMA).Value
                    
                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_osea_sh_nm))
                With .Sheets(SIXP.G_osea_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt

                    
                    'Public Enum E_OSEA_ORDER
                    '    e_osea_order_total = 5
                    '    e_osea_order_on_stock
                    '    e_osea_order_ordered
                    '    e_osea_order_confirmed
                    '    e_osea_order_for_mrd
                    '    e_osea_order_after_mrd
                    '    e_osea_order_open
                    'End Enum
                    
                    .Cells(sixp_w, SIXP.e_osea_order_total).Value = mt.Cells(w, SIXP.e_qmt_osea_Total_osea).Value
                    .Cells(sixp_w, SIXP.e_osea_order_on_stock).Value = mt.Cells(w, SIXP.e_qmt_ON_STOCK_osea).Value
                    .Cells(sixp_w, SIXP.e_osea_order_ordered).Value = mt.Cells(w, SIXP.e_qmt_ORDERED_osea).Value
                    .Cells(sixp_w, SIXP.e_osea_order_confirmed).Value = mt.Cells(w, SIXP.e_qmt_CONFIRMED_osea).Value
                    .Cells(sixp_w, SIXP.e_osea_order_for_mrd).Value = mt.Cells(w, SIXP.e_qmt_FOR_MRD_osea).Value
                    .Cells(sixp_w, SIXP.e_osea_order_after_mrd).Value = mt.Cells(w, SIXP.e_qmt_after_MRD_osea).Value
                    .Cells(sixp_w, SIXP.e_osea_order_open).Value = mt.Cells(w, SIXP.e_qmt_OPEN_osea).Value
                    
                    
                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_totals_sh_nm))
                With .Sheets(SIXP.G_totals_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt

                    
                    'Public Enum e_5p_totals
                    '    e_5p_total = 5
                    '    e_5p_na
                    '    e_5p_itdc
                    '    e_5p_pnoc
                    '    e_5p_fma_eur
                    '    e_5p_fma_osea
                    '    e_5p_ordered
                    '    e_5p_arrived
                    '    e_5p_in_transit
                    '    e_5p_future
                    '    e_5p_ppap_status
                    '    e_5p_no_ppap_status
                    'End Enum
                    
                    
                    ' TOTAL!
                    ' .Cells(sixp_w, SIXP.e_5p_total).Value = 0 ' brak pola w starym quarterze, nalezy dokonac sumy!
                    ' suma sklada sie:
                    ' NA, ITDC, PNOC, FMA EUR, FMA OSEA
                    .Cells(sixp_w, SIXP.e_5p_total).Value = _
                        CLng(mt.Cells(w, SIXP.e_qmt_NA).Value) + _
                        CLng(mt.Cells(w, SIXP.e_qmt_ITDC).Value) + _
                        CLng(mt.Cells(w, SIXP.e_qmt_PNOC_ChartTotal).Value) + _
                        CLng(mt.Cells(w, SIXP.e_qmt_FMAeur).Value) + _
                        CLng(mt.Cells(w, SIXP.e_qmt_FMAosea).Value)

                    .Cells(sixp_w, SIXP.e_5p_na).Value = mt.Cells(w, SIXP.e_qmt_NA).Value
                    .Cells(sixp_w, SIXP.e_5p_itdc).Value = mt.Cells(w, SIXP.e_qmt_ITDC).Value
                    .Cells(sixp_w, SIXP.e_5p_pnoc).Value = mt.Cells(w, SIXP.e_qmt_PNOC_ChartTotal).Value
                    .Cells(sixp_w, SIXP.e_5p_fma_eur).Value = mt.Cells(w, SIXP.e_qmt_FMAeur).Value
                    .Cells(sixp_w, SIXP.e_5p_fma_osea).Value = mt.Cells(w, SIXP.e_qmt_FMAosea).Value
                    .Cells(sixp_w, SIXP.e_5p_ordered).Value = mt.Cells(w, SIXP.e_qmt_ORDERED).Value
                    .Cells(sixp_w, SIXP.e_5p_arrived).Value = mt.Cells(w, SIXP.e_qmt_ARRIVED).Value
                    .Cells(sixp_w, SIXP.e_5p_in_transit).Value = mt.Cells(w, SIXP.e_qmt_Intransit).Value
                    .Cells(sixp_w, SIXP.e_5p_future).Value = 0 ' no field
                    .Cells(sixp_w, SIXP.e_5p_ppap_status).Value = mt.Cells(w, SIXP.e_qmt_PPAP_status_forFMA_part_numbers).Value
                    .Cells(sixp_w, SIXP.e_5p_no_ppap_status).Value = mt.Cells(w, SIXP.e_qmt_not_PPAP_status_for_FMA_part_numbers).Value
                    
                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                SIXP.LoadingFormModule.incLoadingForm
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_resp_sh_nm))
                With .Sheets(SIXP.G_resp_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt
                    'Public Enum E_RESP_ORDER
                    '    e_resp_fma = 5
                    '    e_resp_osea
                    '    e_resp_pem
                    '    e_resp_ppm
                    '    e_resp_sqe
                    'End Enum
                    
                    .Cells(sixp_w, SIXP.e_resp_fma).Value = mt.Cells(w, SIXP.e_qmt_PL).Value
                    .Cells(sixp_w, SIXP.e_resp_osea).Value = mt.Cells(w, SIXP.e_qmt_osea).Value
                    .Cells(sixp_w, SIXP.e_resp_pem).Value = mt.Cells(w, SIXP.e_qmt_pem).Value
                    .Cells(sixp_w, SIXP.e_resp_ppm).Value = mt.Cells(w, SIXP.e_qmt_ppm).Value
                    .Cells(sixp_w, SIXP.e_resp_sqe).Value = mt.Cells(w, SIXP.e_qmt_sqe).Value
                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                SIXP.LoadingFormModule.incLoadingForm
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm))
                With .Sheets(SIXP.G_del_conf_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt
                    
                    
                    'Public Enum E_DEL_CONF_ORDER
                    '    e_del_conf_on_stock = 5
                    '    e_del_conf_edi
                    '    e_del_conf_ho
                    '    e_del_conf_na
                    '    e_del_conf_for_mrd
                    '    e_del_conf_after_mrd
                    '    e_del_conf_for_smrd
                    '    e_del_conf_after_smrd
                    '    e_del_conf_for_twomrd
                    '    e_del_conf_after_twomrd
                    '    e_del_conf_for_twosmrd
                    '    e_del_conf_after_twosmrd
                    '    e_del_conf_for_alt
                    '    e_del_conf_after_alt
                    '    e_del_conf_for_alttwomrd
                    '    e_del_conf_after_alttwomrd
                    '    e_del_conf_for_salttwomrd
                    '    e_del_conf_after_salttwomrd
                    '    e_del_conf_for_oncostmrd
                    '    e_del_conf_after_oncostmrd
                    '    e_del_conf_for_soncostmrd
                    '    e_del_conf_after_soncostmrd
                    '    e_del_conf_open
                    '    e_del_conf_too_late
                    '    e_del_conf_pot_itdc
                    'End Enum
                    
                    .Cells(sixp_w, SIXP.e_del_conf_on_stock).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_edi).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_ho).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_na).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_mrd).Value = mt.Cells(w, SIXP.e_qmt_CONFIRMEDforMRD).Value
                    .Cells(sixp_w, SIXP.e_del_conf_after_mrd).Value = mt.Cells(w, SIXP.e_qmt_ConfafterMRD).Value
                    .Cells(sixp_w, SIXP.e_del_conf_for_smrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_smrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_twomrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_twomrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_twosmrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_twosmrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_alt).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_alt).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_alttwomrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_alttwomrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_salttwomrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_salttwomrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_oncostmrd).Value = mt.Cells(w, SIXP.e_qmt_Oncost).Value
                    .Cells(sixp_w, SIXP.e_del_conf_after_oncostmrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_for_soncostmrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_after_soncostmrd).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_open).Value = mt.Cells(w, SIXP.e_qmt_Open_Chart).Value
                    .Cells(sixp_w, SIXP.e_del_conf_too_late).Value = 0
                    .Cells(sixp_w, SIXP.e_del_conf_pot_itdc).Value = 0
                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                SIXP.LoadingFormModule.incLoadingForm
                SIXP.LoadingFormModule.incLoadingForm
                
                sixp_w = znajdzPierwszyPustyWiersz(ThisWorkbook.Sheets(SIXP.G_xq_sh_nm))
                With .Sheets(SIXP.G_xq_sh_nm)
                    
                    cellsWithLinkProjPltPhaseCW sixp_w, w, .Cells(1, 1).Parent, mt

                    
                    'Public Enum E_XQ_ORDER
                    '    e_xq_comment = 5
                    '    e_xq_ppap_gate
                    '    e_xq_project_type
                    'End Enum
                    
                    .Cells(sixp_w, SIXP.e_xq_comment).Value = mt.Cells(w, SIXP.e_qmt_xq_Comment).Value
                    .Cells(sixp_w, SIXP.e_xq_ppap_gate).Value = mt.Cells(w, SIXP.e_qmt_XQPPAPGate).Value
                    .Cells(sixp_w, SIXP.e_xq_project_type).Value = mt.Cells(w, SIXP.e_qmt_XQProjectType).Value
                    
                    
                End With
                
                
                SIXP.LoadingFormModule.incLoadingForm
                ' back to main
                With .Sheets(SIXP.G_main_sh_nm)

                    ' ----------------------------------------------------------------------------------------
                    
                    For x = SIXP.e_main_last_update_on_order_release_status To SIXP.e_main_last_update_on_resp
                    
                    
                        If x <> SIXP.e_main_last_update_on_open_issues Then
                            .Cells(sixp_w, x).Value = .Cells(sixp_w, SIXP.e_link_cw).Value
                        End If
                    Next x
                    ' ----------------------------------------------------------------------------------------
                    
                    ' reszta danych zostanie uzupelniona potem, gdy dane na innych arkuszach sie juz pojawia
                End With
                
            End With
            
            
            
            ' od teraz r to zasieg danych calego wiersza w starym Q - teraz nalezy przyporzadkowac dane
            ' do nowych wierszy w poszczegolnych arkuszach sixp
        Next w
    End If
End Sub

Private Sub cellsWithLinkProjPltPhaseCW(main_w, w, Sh As Worksheet, mt As Worksheet)

    SIXP.LoadingFormModule.incLoadingForm
    With Sh
        .Cells(main_w, SIXP.e_link_project).Value = mt.Cells(w, SIXP.e_qmt_project).Value
        .Cells(main_w, SIXP.e_link_plt).Value = mt.Cells(w, SIXP.e_qmt_Plant).Value
        .Cells(main_w, SIXP.e_link_faza).Value = mt.Cells(w, SIXP.e_qmt_PHASE).Value
        .Cells(main_w, SIXP.e_link_cw).Value = mt.Cells(w, SIXP.e_qmt_cw).Value ' format YYYYCW
    End With
End Sub


Private Function znajdzPierwszyPustyWiersz(arkusz As Worksheet) As Long

    znajdzPierwszyPustyWiersz = -1
    
    SIXP.LoadingFormModule.incLoadingForm
    Dim r As Range
    Set r = arkusz.Cells(1, 1)
    Do
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    znajdzPierwszyPustyWiersz = r.Row
    
    
    If znajdzPierwszyPustyWiersz = -1 Then
        MsgBox "blad krtyczny w znajdzPierwszyPustyWiersz - program konczy dzialanie!"
        End
    End If
End Function
