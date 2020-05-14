VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewOnePagerHandler"
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

Private kolekcja_linkow As Collection
Private link_raportu As T_Link
Private my_linker As Linker


Private m As Worksheet
Private pcs(1 To 10) As Worksheet

Private one_pager_sheet As Worksheet


Private ppApp As PowerPoint.Application
Private pp As Presentation
Private ppSlide As PowerPoint.Slide
Private myCustomLayout As PowerPoint.CustomLayout




' new table from one pagers - for rep goals
' -----------------------------------------------

Private newTableWrk As Workbook
Private newTableSh As Worksheet
Private newTableRefRange As Range
Private newTableRefRangeGYR As Range
Private newTableRefRangePUS As Range
Private newTabelRefRangeComment As Range
Private newTableRefRangeTot As Range
Private newTableRefRangeDelConf As Range





Public Function get_one_pager_sheet() As Worksheet
    Set get_one_pager_sheet = one_pager_sheet
End Function


Private Sub Class_Initialize()
    Set kolekcja_linkow = New Collection
    
    
    ' procedura z moduly (public sub!)
    clear_one_pager
    
    
    ' heading
    ' 1 == main enum
    Set pcs(1) = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' left side
    ' 2 == order release status enum
    Set pcs(2) = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
    Set pcs(3) = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
    Set pcs(4) = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
    
    ' right side
    ' 5 == recent build plan changes enum
    Set pcs(5) = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
    Set pcs(6) = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
    Set pcs(7) = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
    
    ' bottom
    ' 8 == resp enum
    Set pcs(8) = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
    
    ' second page
    ' 9 == open issues enum
    Set pcs(9) = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
    
    Set pcs(10) = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
    
    
    Set one_pager_sheet = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
    
    Set chart1 = ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm)
    Set chart2 = ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm)
    Set chart3 = ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm)
End Sub

Private Sub Class_Terminate()
    Set kolekcja_linkow = Nothing
End Sub

Public Sub przypisz_kolekcje_linkow(ByRef cl As Collection)
    
    Set kolekcja_linkow = cl
End Sub


Public Sub generuj_raporty(e As E_ONE_PAGERS_INTO)



    ' SIXP.LoadingFormModule -> show zrobiony juz w poprzednim lejerze

    Dim sh As Worksheet, r As Range, rf As Range, pD As Range
    
    inicjuj_power_pointa_jesli_taka_jest_konfiguracja e
    inicjuj_niezalezna_tabele_w_excelu_jesli_jest_taka_konfiguracja e

    Dim iter As Long
    iter = 1
    For Each link_raportu In kolekcja_linkow
    
        
        
        Set r = Nothing
        
        ' cialo generowania iterowanego one pagera
        ' =============================================
        For x = LBound(pcs) To UBound(pcs)
            Set sh = pcs(x)
            sh.Activate
            SIXP.LoadingFormModule.incLoadingForm
            
            ' to jest najprostsza logika dzialania jednak to nie jest tak chocby
            ' dla pcs(2) to jest dla danych order release status trzeba sprawdzic nawet wszystkie fazy dla wybranego projektu
            ' wiec odrobine smutek - gdyz trzeba rozszerzyc dzialanie mocno...
            
            
            
            ' patrzymy tylko w main == x == 1
            If x = 1 Then Set r = link_raportu.znajdz_siebie_w_arkuszu(sh)
            
            If Not r Is Nothing Then
                ' r.Select
                ' MsgBox sh.Name & " " & r & " " & r.Offset(0, 1) & " " & r.Offset(0, 3)
                
                ' niech narazie w ramach wersji testowej robia sie aktywacje oraz selekcje
                ' --------------------------------------------------------------------------
                ' --------------------------------------------------------------------------
                
                If x = SIXP.e_pcs_main Then
                    wypelnienie_statusu_z_main r, link_raportu
                ElseIf x = SIXP.e_pcs_order_release_status Then
                    wypelnienie_pcs_order_release_status rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_pnoc Then
                    wypelnienie_pcs_pnoc rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_osea Then
                    wypelnienie_pcs_osea rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_recent_build_plan_changes Then
                    wypelnienie_pcs_recent_build_plan_changes rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_totals Then
                    wypelnienie_pcs_totals rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_delivery_confirmation Then
                    wypelnienie_pcs_del_conf rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_resp Then
                    wypelnienie_pcs_resp rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_open_issues Then
                    wypelnienie_pcs_open_issues rf, pD, sh, link_raportu
                ElseIf x = SIXP.e_pcs_xq Then
                    wypelnienie_pcs_xq rf, pD, sh, link_raportu
                End If
                
                
                ' --------------------------------------------------------------------------
                ' --------------------------------------------------------------------------
            Else
                ' w main w ogole nie odnalazlem tego ajtemu
                MsgBox "brak danych!"
                End
            End If
            
            
        Next x
        ' =============================================
        
        ' wczesniej ta funkcjonalnosc wykonywana byla przed zmiana danych zatem od samego poczatku zle pokazywalo
        ' jakby offset -1
        With link_raportu
            addHeader .plt & " " & .project & " " & .faza & " " & .cw
        End With
        
        
        przesortuj_raz_jeszcze_order_release_status_w_zgodzie_z_phase_list link_raportu
        szary_na_tym_raporcie link_raportu
        
        ' tutaj koniec iteracji dla calego linku odnoszacego sie do projektu
        ' w tym miejscu arkusz one pager jest juz uzupelniony zatem bez cisnien
        ' mozna dokonac wyciagniecia tego projektu do niezaleznego excela albo do kolejnego slajdu w power poincie
        raport_dla_tego_projektu_gotowy_wyrzuc_go_do_swiata e, CLng(iter), link_raportu
        
        iter = iter + 1
    Next link_raportu
    
    
    pcs(1).Activate
    
    If e = E_ONE_PAGERS_INTO_NEW_TABLE Then
        
        
        ' this section is about filtering grey fields
        ' -------------------------------------------------' -------------------------------------------------
        
        ' newTableWrk
        
        ' za kazdym razem trzeba to zmienic jesli bedzie zmiana ukladu danych!
        
        ' labelki pionowo
        newTableSh.Cells(1, 1).RowHeight = 120

        With newTableSh.Cells(1, 1).EntireRow
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        newTableSh.Range("A1").EntireColumn.ColumnWidth = 8
        newTableSh.Range("B1").EntireColumn.ColumnWidth = 16
        newTableSh.Range("C1").EntireColumn.ColumnWidth = 5
        newTableSh.Range("D1").EntireColumn.ColumnWidth = 7
        newTableSh.Range("E1").EntireColumn.ColumnWidth = 8
        newTableSh.Range("F1").EntireColumn.ColumnWidth = 4
        newTableSh.Range("G1").EntireColumn.ColumnWidth = 6
        newTableSh.Range("H1").EntireColumn.ColumnWidth = 7
        newTableSh.Range("I1").EntireColumn.ColumnWidth = 7
        newTableSh.Range("J1").EntireColumn.ColumnWidth = 7
        newTableSh.Range("K1").EntireColumn.ColumnWidth = 4
        newTableSh.Range("L1").EntireColumn.ColumnWidth = 7
        newTableSh.Range("M1").EntireColumn.ColumnWidth = 7
        
        newTableSh.Range("N1").EntireColumn.ColumnWidth = 3
        
        newTableSh.Range("O1:R1").EntireColumn.ColumnWidth = 4
        
        newTableSh.Range("S1").EntireColumn.ColumnWidth = 6
        
        newTableSh.Range("T1").EntireColumn.ColumnWidth = 6
        newTableSh.Range("U1").EntireColumn.ColumnWidth = 6
        newTableSh.Range("V1").EntireColumn.ColumnWidth = 6
        
        newTableSh.Range("W1").EntireColumn.ColumnWidth = 7
        
        ' newTableSh.Range("A1").AutoFilter Field:=2, Criteria1:=RGB(200, 200, 200), Operator:=xlFilterCellColor
        newTableSh.Range("A1").AutoFilter
        
        newTableSh.Activate
        ActiveWindow.Zoom = 110
        
        
        ' -------------------------------------------------' -------------------------------------------------
    End If
    
    
    clearNewOnePagerWorksheet
    
    
End Sub

Private Sub szary_na_tym_raporcie(l As T_Link)
    
    Dim orsRange As Range
    Set orsRange = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A5:A12")
    
    Dim r As Range
    
    
    For Each r In orsRange
    
        If CStr(UCase(r.Value)) = CStr(UCase(l.faza)) Then
            r.Parent.Range(r, r.Offset(0, 7)).Interior.Color = RGB(200, 200, 200)
        Else
            r.Parent.Range(r, r.Offset(0, 7)).Interior.Color = RGB(255, 255, 255)
        End If
    Next r
    
End Sub

Private Sub przesortuj_raz_jeszcze_order_release_status_w_zgodzie_z_phase_list(l As T_Link)
    
    Dim orsRange As Range
    Set orsRange = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A5:A12")
    
    
    Dim orderFromPhaseListR1 As Integer, orderFromPhaseListR2 As Integer
    
    
    Dim tmp As String
    
    
    Dim r1 As Range, r2 As Range, ir As Range
    For Each ir In orsRange
        If Trim(ir) <> "" Then
        
            If Trim(ir.Offset(1, 0)) <> "" Then
        
                Set r1 = ir
                Set r2 = ir.Offset(1, 0)
                
                orderFromPhaseListR1 = getOrderFromPhaseList(r1)
                orderFromPhaseListR2 = getOrderFromPhaseList(r2)
                
                
                ' jest odwrocona logika!
                If orderFromPhaseListR1 < orderFromPhaseListR2 Then
                
                    For x = 1 To 8
                        tmp = CStr(ir.Offset(0, x - 1).Value)
                        r1.Offset(0, x - 1).Value = CStr(r2.Offset(0, x - 1).Value)
                        r2.Offset(0, x - 1).Value = CStr(tmp)
                    Next x
                    
                    
                    przesortuj_raz_jeszcze_order_release_status_w_zgodzie_z_phase_list l
                    
                End If
                
            Else
                Exit For
            End If
            
        End If
    Next ir
End Sub

Private Function getOrderFromPhaseList(r As Range) As Integer

    getOrderFromPhaseList = 100
    
    Dim plr As Range
    Set plr = ThisWorkbook.Sheets(SIXP.G_PHASE_LIST_SH_NM).Range("A2")
    Set plr = plr.Parent.Range(plr, plr.End(xlDown))
    
    Dim iplr As Range
    For Each iplr In plr
        
        If UCase(Trim(CStr(iplr.Offset(0, 1)))) = UCase(Trim(CStr(r))) Then
            getOrderFromPhaseList = Int(iplr)
            Exit Function
        End If
    Next iplr
    
    
End Function

Private Function assign_qty_to_on_stock(mtl As T_Link) As Long
    assign_qty_to_on_stock = CLng(common_assign_del_conf_qty(mtl, e_del_conf_on_stock))
End Function

Private Function assign_qty_to_edi(mtl As T_Link) As Long
    assign_qty_to_edi = CLng(common_assign_del_conf_qty(mtl, e_del_conf_edi))
End Function

Private Function assign_qty_to_h_slash_o(mtl As T_Link) As Long
    assign_qty_to_h_slash_o = CLng(common_assign_del_conf_qty(mtl, e_del_conf_ho))
End Function

Private Function assign_qty_to_n_slash_a(mtl As T_Link) As Long
    assign_qty_to_n_slash_a = CLng(common_assign_del_conf_qty(mtl, e_del_conf_na))
End Function

Private Function assign_qty_to_mrd_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_mrd))
End Function

Private Function assign_qty_to_mrd_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_mrd))
End Function

Private Function assign_qty_to_mrd_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_mrd))
End Function

Private Function assign_qty_to_mrd_alt_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_alt))
End Function

Private Function assign_qty_to_mrd_alt_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_alt))
End Function

Private Function assign_qty_to_mrd_alt_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_alt))
End Function



Private Function assign_qty_to_mrd_staggered_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_smrd))
End Function

Private Function assign_qty_to_mrd_staggered_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_smrd))
End Function


Private Function assign_qty_to_mrd_staggered_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_smrd))
End Function



' obsolete
' --------------------------------------------------------------------------------------------------------------------
Private Function assign_qty_to_mrd_staggered_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_twosmrd))
End Function

Private Function assign_qty_to_mrd_staggered_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_twosmrd))
End Function
' --------------------------------------------------------------------------------------------------------------------



' new
' --------------------------------------------------------------------------------------------------------------------
Private Function assign_qty_to_mrd_staggered_oncost_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_oncost_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_soncostmrd))
End Function

Private Function assign_qty_to_mrd_staggered_oncost_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_oncost_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_soncostmrd))
End Function

Private Function assign_qty_to_mrd_staggered_oncost_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_oncost_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_soncostmrd))
End Function




Private Function assign_qty_to_mrd_oncost_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_oncost_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_oncostmrd))
End Function

Private Function assign_qty_to_mrd_oncost_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_oncost_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_oncostmrd))
End Function

Private Function assign_qty_to_mrd_oncost_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_oncost_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_oncostmrd))
End Function





Private Function assign_qty_to_mrd_alt_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_alttwomrd))
End Function

Private Function assign_qty_to_mrd_alt_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_alttwomrd))
End Function

Private Function assign_qty_to_mrd_alt_two_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_two_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_alttwomrd))
End Function




Private Function assign_qty_to_mrd_staggered_alt_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_alt_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_salttwomrd))
End Function

Private Function assign_qty_to_mrd_staggered_alt_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_alt_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_salttwomrd))
End Function

Private Function assign_qty_to_mrd_staggered_alt_two_after_red(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_alt_two_after_red = CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_salttwomrd))
End Function
' --------------------------------------------------------------------------------------------------------------------




Private Function assign_qty_to_mrd_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_twomrd))
End Function

Private Function assign_qty_to_mrd_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_twomrd))
End Function



Private Function assign_qty_to_yellow_open(mtl As T_Link) As Long
    assign_qty_to_yellow_open = CLng(common_assign_del_conf_qty(mtl, e_del_conf_yellow_open))
End Function

Private Function assign_qty_to_red_open(mtl As T_Link) As Long
    assign_qty_to_red_open = CLng(common_assign_del_conf_qty(mtl, e_del_conf_open))
End Function

Private Function assign_qty_to_pot_itdc(mtl As T_Link) As Long
    assign_qty_to_pot_itdc = CLng(common_assign_del_conf_qty(mtl, e_del_conf_pot_itdc))
End Function

Private Function assign_qty_to_too_late(mtl As T_Link) As Long
    ' T32 - too late open / alias undef - tak bylo wczesniej
    assign_qty_to_too_late = CLng(common_assign_del_conf_qty(mtl, e_del_conf_too_late))
End Function




Private Function common_assign_del_conf_qty(mtl As T_Link, e As E_DEL_CONF_ORDER) As Long
    common_assign_qty = 0
    
    
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    
    Dim del_confs(0 To 512) As Range, del_conf_rf As Range
    
    Set del_conf_rf = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(2, 1)
    
    xx = 0
    Do
        If Trim(del_conf_rf) = "" Then
            Exit Do
        End If
        
        
        
        If Trim(del_conf_rf) = Trim(mtl.project) Then
        
            If Trim(del_conf_rf.Offset(0, 1)) = Trim(mtl.plt) Then
            
                If Trim(del_conf_rf.Offset(0, 2)) = Trim(mtl.faza) Then
                    ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                    ' o ktory sami poprosilismy :)
                    ' dzieki temu ze wykorzystujemy zapis yyyycw
                    ' mozemy korzystac ze zwyklej arytmetyki >< :)
                    If CLng(del_conf_rf.Offset(0, 3)) <= CLng(mtl.cw) Then
                        Set del_confs(xx) = del_conf_rf
                        xx = xx + 1
                    End If
                End If
            End If
            
        End If
        
        Set del_conf_rf = del_conf_rf.Offset(1, 0)
    Loop Until Trim(del_conf_rf) = ""
    
    bubble_sort del_confs, 3, E_DESCENDING
    
    
    ' ok dane najmlodsze na samej "gorze" od pierwszego itemu
    ' =========================================================
    i = 0
    ' poczatek danych order release na one pagerze
    'Dim pd As Range
    
    xxx = LBound(del_confs)
    Set del_conf_rf = del_confs(xxx)
    
    
    If del_conf_rf Is Nothing Then
    Else
        If IsNumeric(del_conf_rf.Offset(0, e - 1)) Then
            common_assign_del_conf_qty = CLng(del_conf_rf.Offset(0, e - 1))
        Else
            common_assign_del_conf_qty = 0
        End If
        
    End If
        

End Function


Private Function warunek_czy_jest_dana_faza_sie_nie_pojawila(zasieg_do_przejrzenia As Range, faza As Range) As Boolean
    warunek_czy_jest_dana_faza_sie_nie_pojawila = True
    
    Dim ir As Range
    For Each ir In zasieg_do_przejrzenia
        If CStr(ir) = CStr(faza) Then
            warunek_czy_jest_dana_faza_sie_nie_pojawila = False
        End If
    Next ir
End Function


Private Sub raport_dla_tego_projektu_gotowy_wyrzuc_go_do_swiata(e As E_ONE_PAGERS_INTO, iter As Long, ByRef l As T_Link)
    
    ' jestesmy w subie, ktory wyrzuca gotowe dane albo do excela niezaleznego albo do power pointa
    Dim myOwnIter As Long
    Dim source_sh As Worksheet, w As Workbook
    Dim x As Integer
    
    Dim ir As Range
    
    
    If e = E_ONE_PAGERS_INTO_NEW_TABLE Then
    
    
        GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = False
    
    
        ' brand new - to create new table from all data from 6p
        ' for sure will not make any new slides!
        ' but still limitation because we are working on data put in one pager sheet
        ' --------------------------------------------------------------------------
        ' newTableWrk
        ' newTableSh
        
        newTableSh.Activate
        
        ' my source
        Set source_sh = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
        
        'If newTableRefRange.Row = 1 Then
        '    ' fresh just started!
        '    putSomeLabels newTableRefRange, source_sh
        '    Set newTableRefRange = newTableRefRange.Offset(1, 0)
        '    Set newTableRefRangeGYR = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR)
        '    Set newTableRefRangePUS = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS)
        'Else
        '    Set newTableRefRange = newTableRefRange.Offset(8, 0)
        '    Set newTableRefRangeGYR = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR)
        '    Set newTableRefRangePUS = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS)
        'End If
        
        If newTableRefRange.Row = 1 Then
            putSomeLabels newTableRefRange, source_sh
        End If
        
        Set newTableRefRange = newTableRefRange.Offset(1, 0)
        Set newTableRefRangeGYR = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR)
        Set newTableRefRangePUS = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS)
        

        
        ' Debug.Print newTableRefRange.Address & " " & newTableRefRange.Parent.name
        
        
        ' magia kopiowania gorny-lewy rog one pagera
        ' ----------------------------------------------------------
        Dim srcRange As Range, tmpForNewTableRefRnage As Range
        Set srcRange = source_sh.Range("A5:H12")
        
        srcRange.Copy newTableRefRange
        ' ----------------------------------------------------------
        
        Set tmpForNewTableRefRnage = newTableRefRange.Offset(-1, 0)
        
        ' checkIfGreyAndRemoveWithShiftUp usunac
        For x = 7 To 0 Step -1
            checkIfGreyAndRemoveWithShiftUp _
                newTableRefRange.Parent.Cells(newTableRefRange.Row + x, newTableRefRange.Column)
        Next x
        
        
        ' poniewaz przez to usuwanie referencje sie stracily
        ' i trzeba jeszcze raz!
        Set newTableRefRange = tmpForNewTableRefRnage.Offset(1, 0)
        Set newTableRefRangeGYR = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR)
        Set newTableRefRangePUS = newTableRefRange.Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS)
        
        ' koniec z bez sensu petla!
        x = 0
        
        tmpArr = Split(CStr(source_sh.Range("b1").Value), ", ")
        
        
        
        
        q = LBound(tmpArr)
        newTableRefRange.Offset(x, -5).Value = tmpArr(q)
        newTableRefRange.Offset(x, -4).Value = tmpArr(q + 1)
        newTableRefRange.Offset(x, -3).Value = tmpArr(q + 2)
        newTableRefRange.Offset(x, -2).Value = tmpArr(q + 3)
    
            
        newTableRefRange.Offset(x, -6).Interior.Color = newTableRefRange.Offset(x, 1).Interior.Color
        newTableRefRange.Offset(x, -5).Interior.Color = newTableRefRange.Offset(x, 1).Interior.Color
        newTableRefRange.Offset(x, -4).Interior.Color = newTableRefRange.Offset(x, 1).Interior.Color
        newTableRefRange.Offset(x, -3).Interior.Color = newTableRefRange.Offset(x, 1).Interior.Color
        newTableRefRange.Offset(x, -2).Interior.Color = newTableRefRange.Offset(x, 1).Interior.Color
        
        
        
        newTableRefRange.Offset(x, -1).Value = source_sh.Range("m1").Value
        
        If UCase(newTableRefRange.Offset(x, -1).Value) = "X" Then
            newTableRefRange.Offset(x, -1).Interior.Color = RGB(255, 0, 0)
        ElseIf UCase(newTableRefRange.Offset(x, -1).Value) = "O" Then
            newTableRefRange.Offset(x, -1).Interior.Color = RGB(0, 255, 0)
        Else
            newTableRefRange.Offset(x, -1).Interior.Color = RGB(255, 255, 0)
        End If
        
        
        
        ' only for specific (grey) line
        If newTableRefRange.Offset(x, 1).Interior.Color = RGB(200, 200, 200) Then
        
        
            ' --------------------------------------------------------------------
            ' pus arrived
            newTableRefRangePUS.Offset(x, 0).Value = source_sh.Range("e17").Value
            ' PUS In transit
            newTableRefRangePUS.Offset(x, 1).Value = source_sh.Range("e18").Value
            ' PUS in future
            newTableRefRangePUS.Offset(x, 2).Value = source_sh.Range("e19").Value
            ' PUS pending
            newTableRefRangePUS.Offset(x, 3).Value = source_sh.Range("e20").Value
            '
            ' --------------------------------------------------------------------
        
            ' RYG
            newTableRefRangeGYR.Offset(x, 1).Value = source_sh.Range("k29").Value
            newTableRefRangeGYR.Offset(x, 2).Value = source_sh.Range("l29").Value
            newTableRefRangeGYR.Offset(x, 3).Value = source_sh.Range("m29").Value
            
            
            ' OPEN ISSUES
            ' --------------------------------------------------------------------
            
            oddelegujTextBoxPodOpenIssuesDlaX newTableRefRangeGYR.Offset(x, 3), source_sh
            
            ' --------------------------------------------------------------------
            
            
            ' total
            newTableRefRangeGYR.Offset(x, 0).Value = 0
            On Error Resume Next
            ' to jest bez pnoc
            ' newTableRefRangeGYR.Offset(0, 0).Value = CLng(source_sh.Range("e16").Value) + CLng(source_sh.Range("f16").Value)
            newTableRefRangeGYR.Offset(x, 0).Value = CLng(source_sh.Range("a16").Value)
            
            
            ' put comments as excel comment
            Set newTabelRefRangeComment = newTableRefRangeGYR.Offset(x, 4)
            
            
            ' RECENT BUILD PLAN CHANGES!
            ' aktualnie kolumna zaraz po GYR - dla danych z czesci 6p gdzie jest last update
            ' ----------------------------------------------------------------------------------------
            oddelegujTextBoxPodRecentBuildPlanChanges newTabelRefRangeComment, ir, source_sh
            ' ----------------------------------------------------------------------------------------
            
            
            Set newTableRefRangeTot = newTableRefRangeGYR
            
            ' TOTALS
            ' ----------------------------------------------------------------------------------------
            oddlegujTextBoxPodTotals newTableRefRangeTot, ir, source_sh
            ' ----------------------------------------------------------------------------------------
            
            
            Set newTableRefRangeDelConf = newTableRefRangeGYR.Offset(0, 1)
            
            ' DELIVERY CONFIRMATION
            ' ----------------------------------------------------------------------------------------
            oddelegujTextBoxPodDeliveryConfirmation newTableRefRangeDelConf, ir, source_sh
            ' ----------------------------------------------------------------------------------------
            
            
            ' resp fma
            newTableRefRange.Offset(x, -6).Value = source_sh.Range("respfma").Value
            
            
            ' NOK
            ' makeSomeLines newTableRefRangeGYR.Parent.Range(newTableRefRangeGYR.Offset(x, 0), newTableRefRangeGYR.Offset(x, 4))
            
            ' all!
            makeSomeLines _
                newTableRefRangeGYR.Parent.Range(newTableRefRangeGYR.Parent.Cells(newTableRefRangeGYR.Row, 1), _
                newTableRefRangeGYR.Offset(x, 6))
        End If
            
        
        GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = True
        
        ' --------------------------------------------------------------------------
    
    
    ElseIf e = E_ONE_PAGERS_INTO_POWER_POINT Then
    
        
        myOwnIter = iter * 2 - 1
    
        DoEvents
        makeSlide CLng(myOwnIter), ppApp, pp, myCustomLayout, l

    ElseIf e = E_ONE_PAGERS_INTO_SEPERATE_EXCELS Then
        
        ' nowy pliczor
        Set w = Workbooks.Add
        
        ' tutaj akursz one pagera
        Set source_sh = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
        source_sh.Copy Before:=w.Sheets(1)
        
        ' pasuje dodac jeszcze arkusze zawierajace zrodla dla
        ' charty 1, 2, 3
        'ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Copy Before:=w.Sheets(1)
        'ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm).Copy Before:=w.Sheets(1)
        'ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm).Copy Before:=w.Sheets(1)
        
        
        Dim obiekty As ChartObjects
        Dim source_range As Range, rep_sh As Worksheet

        
        
        Set rep_sh = w.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
        rep_sh.Activate
        
        
        Dim changeR As Range
        Set changeR = rep_sh.Range("B17:F18")
        
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        
        ' open bp oraz actionable FMA
        Set changeR = rep_sh.Range("B14")
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        Set changeR = rep_sh.Range("E14")
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        
        ' oversea
        Set changeR = rep_sh.Range("A22:G22")
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        ' overall
        Set changeR = rep_sh.Range("A26:G26")
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        Set changeR = rep_sh.Range("A29:C29")
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        Set changeR = rep_sh.Range("E29:F29")
        For Each ir In changeR
            ir.Value = CStr(ir.Value)
        Next ir
        
        
        
        ' plik w to jest nowy workbook w ktorym znajduje sie pojedynczy raport dla jednej "linijki" z 6P
        
        Application.DisplayAlerts = False
        w.SaveAs ThisWorkbook.Path & Application.PathSeparator & _
            CStr(SIXP.G_EXCEL_REPORT_PREFIX & _
                Replace(l.plt, ":", "") & "_" & Replace(l.project, ":", "") & "_" & CStr(l.faza) & "_" & CStr(l.cw)), _
            51, , , , , , xlLocalSessionChanges
        Application.DisplayAlerts = True
        
        
        w.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Activate
        
        
        ' to gdy pracujemy na outlooku
        '.Attachments.Add w.FullName
        ' nie chcemy zamykac niech sobie wisi
        ' w.Close False
    Else
        MsgBox "ta sytuacja nie jest mozliwa - blad krytyczny - makro zatrzymalo sie!"
        End
    End If
End Sub


Private Sub oddelegujTextBoxPodOpenIssuesDlaX(x As Range, src As Worksheet)



        ' sekcja tworzenia textboxa!
        ' //////////////////////////////////////////////////////////////
        
        Dim tbh As TextBoxHandler
        Set tbh = New TextBoxHandler
        
        
        '
        ' ROW 5 is starting position
        '
        '' STATUS 2 kolumny
        'ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("P" & CStr(5 + i)) = _
        '    rf.Offset(0, SIXP.e_open_issues_status - 1)
        '
        '' # of pn 2 kolumny
        'ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("R" & CStr(5 + i)) = _
        '    rf.Offset(0, SIXP.e_open_issues_no_of_pn - 1)
        '
        '' part  / supplier 3 kolumny
        'ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("T" & CStr(5 + i)) = _
        '    rf.Offset(0, SIXP.e_open_issues_part_supplier - 1)
        '
        '' delivery 2 kolumny
        'ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("W" & CStr(5 + i)) = _
        '    rf.Offset(0, SIXP.e_open_issues_delivery - 1)
        '
        '' cmnt 15 kolumn
        'ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("Y" & CStr(5 + i)) = _
        '    rf.Offset(0, SIXP.e_open_issues_comment - 1)
        '
        'i = i + 2
        
        Dim refRow As Long, p As Range, r As Range, txt As String
        
        ' Dim txt As String
        
        Dim t As Range, w As Range, y As Range
        txt = ""
        refRow = 5
        
        
        Dim tbInTextBox As New TableInTextBox
        tbInTextBox.defineLables "St", "#PNs", "Part or Supplier", "Delivery", "Comments"
        tbInTextBox.defineSizes 5, 10, 20, 10, 40
        tbInTextBox.setCols 5
        
        tbInTextBox.tryToAddLineOfTheData "St", "#PNs", "Part or Supplier", "Delivery", "Comments"
        
        '' STATUS 2 kolumny
        ' czyzby to samo co src
        Set p = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("P5")
        Set r = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("R5")
        ' part supplier
        Set t = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("T5")
        ' delivery
        Set w = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("W5")
        ' cmnt
        Set y = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("Y5")
        
        Do
        
            'txt = txt & _
            '    p.Value & " " & _
            '    r.Value & " " & _
            '    t.Value & " " & _
            '    w.Value & " " & _
            '    y.Value & " " & _
            '    chr(10)
            
            tbInTextBox.tryToAddLineOfTheData p.Value, r.Value, t.Value, w.Value, y.Value
            
            
            
            Set p = p.Offset(1, 0)
            Set r = r.Offset(1, 0)
            Set t = t.Offset(1, 0)
            Set w = w.Offset(1, 0)
            Set y = y.Offset(1, 0)
        Loop Until Trim(p) = ""
        
        
'        Debug.Print txt
'        Dim test As New TableInTextBox
'        test.defineLables "jeden", "dwa", "trzy"
'        test.defineSizes 10, 20, 8
'        test.setCols 3
'        test.tryToAddLineOfTheData "jeden", "dwa", "trzy"
'        test.tryToAddLineOfTheData "1", "to jest troche wiecej textu", "to jest troche wiecej textu"
'        test.tryToAddLineOfTheData "1", "to jest troche wiecej textu", 3
'        test.tryToAddLineOfTheData "jeden", "dwa", 3
'        test.tryToAddLineOfTheData "1", "to jest troche wiecej textu", 3
'        test.tryToAddLineOfTheData "1", 4, "to co innego"
'        test.tryToAddLineOfTheData "ewfdweff euhfwef", 2, 3
'
'
'        Dim c As Collection
'        Set c = test.collOfDataLines
'
'        Dim wynik As String
'        wynik = CStr(test.drawThisTable())
'
'        Debug.Print wynik
        
        txt = CStr(tbInTextBox.drawThisTable())
        Debug.Print txt
        
        tbh.addOpenIssueTextBoxForCell x, txt
        
        Set tbh = Nothing
        
        
        ' //////////////////////////////////////////////////////////////

End Sub



Private Sub oddelegujTextBoxPodRecentBuildPlanChanges(newTabelRefRangeComment As Range, ir As Range, source_sh As Worksheet)

    Dim txt As String
    txt = ""

    Dim tbInTextBox As New TableInTextBox
    tbInTextBox.defineLables "Build phase", "#ofVeh", "TBW", "OrderRel/Chge", "Comments"
    tbInTextBox.defineSizes 16, 10, 15, 15, 30
    tbInTextBox.setCols 5
    
    tbInTextBox.tryToAddLineOfTheData "Build phase", "#ofVeh", "TBW", "OrderRel/Chge", "Comments"

    With newTabelRefRangeComment
    


        
        Set ir = source_sh.Range("J4")
        
        Do
            If Trim(ir.Value) <> "" Then
            
'                If ir.Row = 4 Then
'                    txt = txt & ir.Value & "   " & _
'                        ir.Offset(0, 1).Value & "     " & _
'                        ir.Offset(0, 2).Value & "       " & _
'                        ir.Offset(0, 3).Value & "       " & _
'                        ir.Offset(0, 4).Value & chr(10)
'
'                    '.Value = "" & .Value & "  " & _
'                    '   ir.Value & " " & _
'                    '   ir.Offset(0, 1).Value & "   " & _
'                    '   ir.Offset(0, 2).Value & "        " & _
'                    '   ir.Offset(0, 3).Value & "   " & _
'                    '   ir.Offset(0, 4).Value & chr(10)
'
'                Else
'
'                    txt = txt & ir.Value & "                " & _
'                        ir.Offset(0, 1).Value & "               " & _
'                        ir.Offset(0, 2).Value & "                " & _
'                        ir.Offset(0, 3).Value & "           " & _
'                        ir.Offset(0, 4).Value & chr(10)
'
'
'
'                    '.Value = "" & .Value & " " & _
'                    '   ir.Value & "                     " & _
'                    '   ir.Offset(0, 1).Value & "           " & _
'                    '   ir.Offset(0, 2).Value & "           " & _
'                    '   ir.Offset(0, 3).Value & "                     " & _
'                    '   ir.Offset(0, 4).Value & chr(10)
'                End If
'
'
'

                If ir.Row = 4 Then
                    ' no operation
                    ' =====================
                Else
                
                    tbInTextBox.tryToAddLineOfTheData _
                        ir.Value, ir.Offset(0, 1).Value, ir.Offset(0, 2).Value, _
                        ir.Offset(0, 3).Value, ir.Offset(0, 4).Value
                End If

            End If

            
        
            Set ir = ir.Offset(1, 0)
        Loop Until Trim(ir) = ""
        
        .Value = "X"
        .Font.Size = 12
        .WrapText = True
        
        
        
        'Range("W2").Comment.Text Text:= _
        '    "Build phase # of Veh. TBW Order release / change Comments" & chr(10) & "VRS1 6  201947 Advanceemnt from VRS2 to VRS1" & chr(10) & ""
        'Selection.ShapeRange.IncrementLeft -352.5
        'Selection.ShapeRange.IncrementTop 54.75
        
        
        
        ' sekcja tworzenia standardowego komentarza
        ' =====================================================================
        ' =====================================================================
        ' =====================================================================
        'If .Comment Is Nothing Then
        '
        '    .Comment.Shape.TextFrame.Characters.Font.Size = 7
        '    .AddComment txt
        '    .Comment.Shape.TextFrame.Characters.Font.Size = 7
        '    .Comment.Shape.TextFrame.AutoSize = True
        '    .Comment.Shape.TextFrame.Characters.Font.Size = 7
        '    '.Comment.Shape.IncrementLeft -200
        '    '.Comment.Shape.IncrementTop 40 ' NOK
        '
        '    ' super custom!
        '    ' =========================================================
        '    ' superCustomChangeCommentPosition newTabelRefRangeComment ' NOK!
        '    ' =========================================================
        '
        '    ' .Comment.Shape.Left = .Comment.Shape.Left - 200
        'Else
        '    .Comment.Text CStr(txt)
        '    .Comment.Shape.TextFrame.Characters.Font.Size = 7
        '    .Comment.Shape.TextFrame.AutoSize = True
        '    .Comment.Shape.TextFrame.Characters.Font.Size = 7
        '    '.Comment.Shape.IncrementLeft -200
        '    '.Comment.Shape.IncrementTop 40 ' NOK
        'End If
        ' =====================================================================
        ' =====================================================================
        ' =====================================================================
        
        
        
        
        ' sekcja tworzenia textboxa!
        ' //////////////////////////////////////////////////////////////
        
        ' !
        txt = CStr(tbInTextBox.drawThisTable())
        ' !
        
        Dim tbh As TextBoxHandler
        Set tbh = New TextBoxHandler
        
        tbh.addBasicTextBoxForCell newTabelRefRangeComment, txt
        
        
        Set tbh = Nothing
        
        
        ' //////////////////////////////////////////////////////////////
        
        
    End With
    
    
End Sub



Private Sub oddlegujTextBoxPodTotals(newTabelRefRangeComment As Range, ir As Range, source_sh As Worksheet)

    Dim txt As String
    txt = ""

    Dim tbInTextBox As New TableInTextBox
    tbInTextBox.defineLables CStr(source_sh.Range("A15").Value), _
        CStr(source_sh.Range("A16").Value)
    tbInTextBox.defineSizes 12, 12
    tbInTextBox.setCols 2
    
    ' tbInTextBox.tryToAddLineOfTheData

    With newTabelRefRangeComment
        
        ' tot
        Set ir = source_sh.Range("A16")
        tbInTextBox.tryToAddLineOfTheData ir.Offset(-1, 0), ir
        
        
        ' NA
        Set ir = source_sh.Range("B16")
        tbInTextBox.tryToAddLineOfTheData ir.Offset(-1, 0), ir
        
        ' ITDC
        Set ir = source_sh.Range("C16")
        tbInTextBox.tryToAddLineOfTheData ir.Offset(-1, 0), ir
        
        ' PNOC
        Set ir = source_sh.Range("D16")
        tbInTextBox.tryToAddLineOfTheData ir.Offset(-1, 0), ir
        
        
        ' OSEA
        Set ir = source_sh.Range("E16")
        tbInTextBox.tryToAddLineOfTheData ir.Offset(-1, 0), ir
        
        ' EUR
        Set ir = source_sh.Range("F16")
        tbInTextBox.tryToAddLineOfTheData ir.Offset(-1, 0), ir
        
        
        
        
        
        
        ' sekcja tworzenia textboxa!
        ' //////////////////////////////////////////////////////////////
        
        ' !
        txt = CStr(tbInTextBox.drawThisTable())
        ' !
        
        Dim tbh As TextBoxHandler
        Set tbh = New TextBoxHandler
        
        tbh.addBasicTextBoxForCell newTabelRefRangeComment, txt
        
        
        Set tbh = Nothing
        
        
        ' //////////////////////////////////////////////////////////////
        
        
    End With
    
    
End Sub


Private Sub oddelegujTextBoxPodDeliveryConfirmation(newTabelRefRangeComment As Range, ir As Range, source_sh As Worksheet)



    Dim txt As String, x As Variant
    txt = ""

    Dim tbInTextBox As New TableInTextBox
    tbInTextBox.defineLables "DC", _
        CStr(source_sh.Range("K15").Value), _
        CStr(source_sh.Range("L15").Value), _
        CStr(source_sh.Range("M15").Value)
        
    tbInTextBox.defineSizes 20, 8, 8, 8
    tbInTextBox.setCols 4
    
    tbInTextBox.tryToAddLineOfTheData "DC", _
        CStr(source_sh.Range("K15").Value), _
        CStr(source_sh.Range("L15").Value), _
        CStr(source_sh.Range("M15").Value)

    With newTabelRefRangeComment
    
        For x = 16 To 28
            Set ir = source_sh.Range("J" & CStr(x))
            
            If Len("" & Trim(ir.Offset(0, 1).Value) & Trim(ir.Offset(0, 2).Value) & Trim(ir.Offset(0, 3).Value)) > 0 Then
                tbInTextBox.tryToAddLineOfTheData CStr(ir.Value), _
                    CStr(ir.Offset(0, 1).Value), _
                    CStr(ir.Offset(0, 2).Value), _
                    CStr(ir.Offset(0, 3).Value)
            End If
                
        Next x
        
        
        
        
        ' sekcja tworzenia textboxa!
        ' //////////////////////////////////////////////////////////////
        
        ' !
        txt = CStr(tbInTextBox.drawThisTable())
        ' !
        
        Dim tbh As TextBoxHandler
        Set tbh = New TextBoxHandler
        
        tbh.addBasicTextBoxForCell newTabelRefRangeComment, txt
        
        
        Set tbh = Nothing
        
        
        ' //////////////////////////////////////////////////////////////
        
        
    End With

End Sub



Private Sub makeSomeLines(rfr As Range)


    rfr.Borders(xlDiagonalDown).LineStyle = xlNone
    rfr.Borders(xlDiagonalUp).LineStyle = xlNone
    With rfr.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rfr.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rfr.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rfr.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rfr.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rfr.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub


Private Sub putSomeLabels(rfr As Range, opsh As Worksheet)
    
    With rfr
    
        .Offset(0, -6).Value = "RLOG / PPL"
        .Offset(0, -5).Value = "PROJECT"
        .Offset(0, -4).Value = "PLT"
        .Offset(0, -3).Value = "FAZA"
        .Offset(0, -2).Value = "UPDATED ON"
        .Offset(0, -1).Value = "STATUS"
        
        For col = 1 To 8
            .Offset(0, col - 1).Value = opsh.Cells(4, col).Value
        Next col
        
        
        ' -----------------------------------------------------
        
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS).Value = "PUS arrived / on stock"
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS + 1).Value = "PUS In transit"
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS + 2).Value = "PUS in future"
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RPUS + 3).Value = "PUS pending"

        
        ' -----------------------------------------------------
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR).Value = "PNs"
        opsh.Range("K15:M15").Copy .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR + 1)
        .Offset(0, G_KOLUMNA_NEW_TABLE_RGYR + 4).Value = "COMMENTS"
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR + 5).Value = "Extra COMMENTS 1"
        .Offset(0, SIXP.G_KOLUMNA_NEW_TABLE_RGYR + 6).Value = "Extra COMMENTS 2"
        ' -----------------------------------------------------
        
    End With
End Sub



Private Sub makeSlide(ktorySlide As Long, ppApp As PowerPoint.Application, _
    pp As PowerPoint.Presentation, myCustomLayout As PowerPoint.CustomLayout, _
    Optional l As T_Link)

    

    ' pp.PageSetup.SlideSize = ppSlideSizeA4Paper
    ' pp.PageSetup.SlideOrientation = msoOrientationVertical
    'pp.PageSetup.SlideHeight = 297 * 2.88
    'pp.PageSetup.slideWidth = 210 * 2.88
    copyPasteDataFromOnePager 1, CLng(ktorySlide), ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A1:N31"), l
    copyPasteDataFromOnePager 2, CLng(ktorySlide + 1), ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("P3:AC31"), l

    
End Sub


Private Sub copyPasteDataFromOnePager(page As Integer, _
    slideCounter As Long, _
    copiedRange As Range, _
    Optional l As T_Link)
    
    Dim toCopy As Range, shToCopy As Worksheet
    Set ppSlide = pp.Slides.AddSlide(slideCounter, myCustomLayout)
    
    Set shToCopy = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
    shToCopy.DisplayPageBreaks = False
    
    ' copiedRange
    Set toCopy = copiedRange
    toCopy.Parent.DisplayPageBreaks = False
    toCopy.CopyPicture
    
    ppSlide.Shapes.Paste
    
    Dim myShape As PowerPoint.Shape
    Dim heightDiff As Long
    Dim widthDiff As Long
    Set myShape = ppSlide.Shapes(1)
    
    
        
    Do
        With myShape
        
            .LockAspectRatio = msoTrue
            
            If .Height > (pp.PageSetup.SlideHeight) Then
                .Height = pp.PageSetup.SlideHeight
                Exit Do
            ElseIf .Width > pp.PageSetup.slideWidth Then
                .Width = pp.PageSetup.slideWidth
                Exit Do
            End If
            
            
        End With
        
        If Not ((myShape.Height > pp.PageSetup.SlideHeight) Or (myShape.Width > pp.PageSetup.slideWidth)) Then
            ' Exit Do
            
            heightDiff = pp.PageSetup.SlideHeight - myShape.Height
            widthDiff = pp.PageSetup.slideWidth - myShape.Width
            
            
            If widthDiff < heightDiff Then
                myShape.Width = 0.95 * pp.PageSetup.slideWidth
            Else
                myShape.Height = 0.95 * pp.PageSetup.SlideHeight
            End If
            
            myShape.Top = 10
            myShape.Left = 10
            
            
            Exit Do
        End If
    Loop While True
    
    

End Sub



' new in 2020 :o
Private Sub inicjuj_niezalezna_tabele_w_excelu_jesli_jest_taka_konfiguracja(e As E_ONE_PAGERS_INTO)


    If e = E_ONE_PAGERS_INTO_NEW_TABLE Then
        
        
        ' rezygnuje z pomyslu tabeli w nowym arkuszu, poniewaz traci mi sie molizwosc eventu dbl click
        'Set newTableWrk = Workbooks.Add
        '
        'Do
        '    DoEvents
        'Loop While newTableWrk Is Nothing
        
        
        ' aby zachowac konwencje nazewnicza
        Set newTableWrk = ThisWorkbook
        
        
        ' also NOK
        'Application.DisplayAlerts = False
        '
        ' ------------------------------------------------------------
        'On Error Resume Next
        'ThisWorkbook.Sheets("newTable").Delete
        ' ------------------------------------------------------------
        '
        'Application.DisplayAlerts = True
        
        Set newTableSh = newTableWrk.Sheets("newTable")
        Set newTableForWithEvents = newTableSh
        ' newTableSh.name = "newTable"
        
        ' in clear module as public sub
        clearNewTableSheet
        
        
        ' zmiana z 3 na 6 poniewaz rozbilem string z project, plt, faza i updated
        Set newTableRefRange = newTableSh.Cells(1, G_GDZIE_ZACZYNA_SIE_1TB_DLA_NEW_TABLE)
        Set newTableRefRangeGYR = newTableRefRange.Offset(0, 8)
    End If
End Sub


Private Sub inicjuj_power_pointa_jesli_taka_jest_konfiguracja(e As E_ONE_PAGERS_INTO)
    
    If e = E_ONE_PAGERS_INTO_POWER_POINT Then
        Set ppApp = New PowerPoint.Application
        Set pp = ppApp.Presentations.Add
        Set myCustomLayout = pp.Designs(1).SlideMaster.CustomLayouts(7)
    End If
End Sub

Private Function przelicz(slideWidth, imgWidth) As Integer

    przelicz = Int((slideWidth - imgWidth) / 2)
End Function



Private Sub wypelnienie_statusu_z_main(r As Range, tl As T_Link)
    
    Dim pobierzStatus As Range
    Set pobierzStatus = r.Parent.Cells(r.Row, SIXP.e_main_status)
    
    
    'Dim l As Linker
    'Set l = New Linker
    'txt = l.return_full_concated_r_string_comma_seperated(tl)
    With ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
        .Range("B1").Value = CStr(tl.project) + ", " + CStr(tl.plt) + ", " + CStr(tl.faza) + ", " + CStr(tl.cw)
        ' nowe miejsce!
        .Range("M1").Value = CStr(pobierzStatus)
    End With
End Sub

Private Sub wypelnienie_pcs_order_release_status(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' specjalne zachowanie dla order release status
    ' tworzymy tabele nie tylko dla wybranego linku ale rowniez dla innych elementow
    ' dla innych faz tego samego projektu - co nieco komplikuje sprawe - jednak nie jest
    ' to nie do zrobienia
    ' ------------------------------------------------------------------------------------------------
    
    ' arkusz sh jest juz przypisany musimy zaczac od poczatku jesli chodzi o sprawdzenie tego samego projektu
    ' nawet dla roznych dat, ale do wybranej daty i dla roznych faz
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    
    Dim orders_coll(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna order release status
        If Trim(rf) = Trim(tl.project) Then
        
            ' 2 kolumna order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
            
                ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                ' o ktory sami poprosilismy :)
                ' dzieki temu ze wykorzystujemy zapis yyyycw
                ' mozemy korzystac ze zwyklej arytmetyki >< :)
                If CLng(rf.Offset(0, 3)) <= CLng(tl.cw) Then
                    Set orders_coll(xx) = rf
                    xx = xx + 1
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    
    
    ' bubble sorting po dacie trzeciej kolumny?
    ' tutaj do rozwazenia wykorzystania jednak nie sortowania prostego
    ' jednak wykorzystania listy PHASE_LSIT
    bubble_sort orders_coll, 3, E_DESCENDING
    ' phase_list_sort orders_coll
    
    
    ' ok dane najmlodsze na samej "gorze" od pierwszego itemu
    ' =========================================================
    i = 0
    ' poczatek danych order release na one pagerze
    'Dim pd As Range
    
    
    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A5:H12").Value = ""
    
    Set pD = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A5")
    For xxx = LBound(orders_coll) To UBound(orders_coll)
    
    
        Set rf = orders_coll(xxx)
        
        If Not rf Is Nothing Then
    
            ' warunek czy juz nie przekorczylismy wielkosci tabelki
            If i < SIXP.G_ORDER_RELEASE_STATUS_LIMIT Then
            
            
                ' ten warunek sprawdza, czy przypadkiem juz jeden update fazy sie nie pojawil
                ' koniec koncow z racji tego ze wysortowalem tablice orderow
                ' nie trzeba juz korzystac z phase listy ktora byla poprzednio potrzebna
                ' w staym quarterze
                ' dodatkowo nie jest potrzebny exit for poniewaz pomimo duplikatu chcemy dalej przegldac dane z tabeli
                If warunek_czy_jest_dana_faza_sie_nie_pojawila(ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A5:A12"), rf.Offset(0, SIXP.e_link_faza - 1)) Then
                    
                    ' PHASE
                    pD.Offset(i, 0) = rf.Offset(0, SIXP.e_link_faza - 1)
                    
                    ' MRD
                    pD.Offset(i, 1) = rf.Offset(0, SIXP.e_order_release_mrd - 1)
                    
                    ' Build
                    pD.Offset(i, 2) = rf.Offset(0, SIXP.e_order_release_build - 1)
                    
                    ' bom freeze
                    pD.Offset(i, 3) = rf.Offset(0, SIXP.e_order_release_bom_freeze - 1)
                    
                    ' # of veh
                    pD.Offset(i, 4) = rf.Offset(0, SIXP.e_order_release_no_of_veh - 1)
                    
                    ' orders due
                    pD.Offset(i, 5) = rf.Offset(0, SIXP.e_order_release_orders_due - 1)
                    
                    ' released
                    pD.Offset(i, 6) = rf.Offset(0, SIXP.e_order_release_released - 1)
                    
                    ' weeks delay
                    pD.Offset(i, 7) = rf.Offset(0, SIXP.e_order_release_weeks_delay - 1)
                    
                    
                    i = i + 1
                End If
            Else
                Exit For
            End If
        Else
            Exit For
        End If
            
    Next xxx
    ' =========================================================
    
    
    ' ------------------------------------------------------------------------------------------------
End Sub

Private Sub wypelnienie_pcs_pnoc(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' te dane sa pod wykres wraz z danymi historycznymi
    ' - zatem patrzymy na ten sam projekt z perspektywy
    ' roznych CW
    ' ------------------------------------------------------------------------------------------------
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim fst_chart(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    ' 1.
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna proj
        If Trim(rf) = Trim(tl.project) Then
            ' 2 kolumna plt order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
                ' 3 kolumna faza
                If Trim(rf.Offset(0, 2)) = Trim(tl.faza) Then
            
                    ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                    ' o ktory sami poprosilismy :)
                    ' dzieki temu ze wykorzystujemy zapis yyyycw
                    ' mozemy korzystac ze zwyklej arytmetyki >< :)
                    If CLng(rf.Offset(0, 3)) <= CLng(tl.cw) Then
                        Set fst_chart(xx) = rf
                        xx = xx + 1
                    End If
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    ' 2.
    bubble_sort fst_chart, 3, E_DESCENDING
    
    
    ' 3.
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    '''
    ''
    '
    
    i = 0
    ' poczatek danych order release na one pagerze
    ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Range("D6:F12").Clear
    ' Dim pd As Range
    Set pD = ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Range("D8")
    ' recent_bp_changes()
    For xxx = LBound(fst_chart) To UBound(fst_chart)
    
    
        Set rf = fst_chart(xxx)
        
        If Not rf Is Nothing Then
        
        
            If ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Range("D6").Value = "" Then
            
                ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Range("D6").Value = _
                    rf.Offset(0, SIXP.e_cont_pnoc_chart_open_bp - 1)
                ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Range("E6").Value = _
                    rf.Offset(0, SIXP.e_cont_pnoc_chart_actionable_fma - 1)
                    
                ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Range("F6").Value = _
                    rf.Offset(0, SIXP.e_link_cw - 1)
            End If
    
            ' warunek czy juz nie przekorczylismy wielkosci tabelki
            ' dzielenie przez dwa poniewaz sa mergi i w sumie przeskakuje na co drugi wiersz
            ' sumarycznie contentu jest 4 jednak wizualnie zaafektowane jest 8 wierszy arkusza
            If i < SIXP.G_RECENT_BUILD_PLAN_CHANGES_LIMIT Then
                 
                ' CONTRACTED
                pD.Offset(i, 0) = rf.Offset(0, SIXP.e_cont_pnoc_chart_contracted - 1)
                pD.Offset(i, 1) = rf.Offset(0, SIXP.e_cont_pnoc_chart_pnoc - 1)
                pD.Offset(i, 2) = rf.Offset(0, SIXP.e_link_cw - 1)
                
                i = i + 1
    
            Else
                Exit For
            End If
        Else
            Exit For
        End If
            
    Next xxx
    
    '
    ''
    '''
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    
    ' ------------------------------------------------------------------------------------------------
End Sub

Private Sub wypelnienie_pcs_osea(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim sec_chart(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    ' 1.
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna proj
        If Trim(rf) = Trim(tl.project) Then
            ' 2 kolumna plt order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
                ' 3 kolumna faza
                If Trim(rf.Offset(0, 2)) = Trim(tl.faza) Then
            
                    ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                    ' o ktory sami poprosilismy :)
                    ' dzieki temu ze wykorzystujemy zapis yyyycw
                    ' mozemy korzystac ze zwyklej arytmetyki >< :)
                    If CLng(rf.Offset(0, 3)) <= CLng(tl.cw) Then
                        Set sec_chart(xx) = rf
                        xx = xx + 1
                    End If
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    ' 2.
    bubble_sort sec_chart, 3, E_DESCENDING
    
    
    ' 3.
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    '''
    ''
    '
    
    i = 0
    ' poczatek danych order release na one pagerze
    ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm).Range("C6:I6").Clear
    ' Dim pd As Range
    Set pD = ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm).Range("C6")
    ' recent_bp_changes()
    For xxx = LBound(sec_chart) To UBound(sec_chart)
    
    
        Set rf = sec_chart(xxx)
        
        If Not rf Is Nothing Then
        
        
            If pD.Value = "" Then
            
                pD.Value = rf.Offset(0, SIXP.e_osea_order_total - 1)
                pD.Offset(0, 1).Value = rf.Offset(0, SIXP.e_osea_order_on_stock - 1)
                pD.Offset(0, 2).Value = rf.Offset(0, SIXP.e_osea_order_ordered - 1)
                pD.Offset(0, 3).Value = rf.Offset(0, SIXP.e_osea_order_confirmed - 1)
                pD.Offset(0, 4).Value = rf.Offset(0, SIXP.e_osea_order_for_mrd - 1)
                pD.Offset(0, 5).Value = rf.Offset(0, SIXP.e_osea_order_after_mrd - 1)
                pD.Offset(0, 6).Value = rf.Offset(0, SIXP.e_osea_order_open - 1)
                    
            ElseIf pD.Value <> "" Then
                Exit For
            End If
    
            
        Else
            Exit For
        End If
            
    Next xxx
    
    '
    ''
    '''
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    
    ' ------------------------------------------------------------------------------------------------
End Sub


Private Sub wypelnienie_pcs_recent_build_plan_changes(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim recent_bp_changes(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    
    ' 1.
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna proj
        If Trim(rf) = Trim(tl.project) Then
            ' 2 kolumna plt order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
                ' 3 kolumna faza
                If Trim(rf.Offset(0, 2)) = Trim(tl.faza) Then
            
                    ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                    ' o ktory sami poprosilismy :)
                    ' dzieki temu ze wykorzystujemy zapis yyyycw
                    ' mozemy korzystac ze zwyklej arytmetyki >< :)
                    If CLng(rf.Offset(0, 3)) <= CLng(tl.cw) Then
                        Set recent_bp_changes(xx) = rf
                        xx = xx + 1
                    End If
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    ' 2.
    bubble_sort recent_bp_changes, 3, E_DESCENDING
    
    
    ' 3.
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    '''
    ''
    '
    
    i = 0
    ' poczatek danych order release na one pagerze
    ' Dim pd As Range
    
    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("J5:N11").Value = ""
    
    Set pD = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("J5")
    ' recent_bp_changes()
    For xxx = LBound(recent_bp_changes) To UBound(recent_bp_changes)
    
    
        Set rf = recent_bp_changes(xxx)
        
        If Not rf Is Nothing Then
    
            ' warunek czy juz nie przekorczylismy wielkosci tabelki
            ' dzielenie przez dwa poniewaz sa mergi i w sumie przeskakuje na co drugi wiersz
            ' sumarycznie contentu jest 4 jednak wizualnie zaafektowane jest 8 wierszy arkusza
            If Int(i / 2#) < SIXP.G_RECENT_BUILD_PLAN_CHANGES_LIMIT Then
            
            
                If isThereAnyContent(rf) And checkIfNoDuplicates(rf, ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("J5")) Then
                     
                    ' PHASE
                    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("J" & CStr(5 + i)) = _
                        rf.Offset(0, SIXP.e_link_faza - 1)
                    
                    ' # of veh
                    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("K" & CStr(5 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_no_of_veh - 1)
                    
                    ' TBW
                    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("L" & CStr(5 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_tbw - 1)
                    
                    ' order release change
                    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("M" & CStr(5 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1)
                    
                    ' cmnt
                    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("N" & CStr(5 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_comment - 1)
                    
                    i = i + 2
                End If
    
            Else
                Exit For
            End If
        Else
            Exit For
        End If
            
    Next xxx
    
    '
    ''
    '''
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    
    ' ------------------------------------------------------------------------------------------------
End Sub


Private Function isThereAnyContent(rf As Range) As Boolean
    isThereAnyContent = False
    
    dlugoscStr = 0
    dlugoscStr = dlugoscStr + Len(Trim(rf.Offset(0, SIXP.e_recent_bp_ch_no_of_veh - 1)))
    dlugoscStr = dlugoscStr + Len(Trim(rf.Offset(0, SIXP.e_recent_bp_ch_tbw - 1)))
    dlugoscStr = dlugoscStr + Len(Trim(rf.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1)))
    dlugoscStr = dlugoscStr + Len(Trim(rf.Offset(0, SIXP.e_recent_bp_ch_comment - 1)))
    
    
    If dlugoscStr > 0 Then
        isThereAnyContent = True
    End If
    
End Function


Private Function checkIfNoDuplicates(rf As Range, onePager_Range_RBPC As Range) As Boolean


    checkIfNoDuplicates = True

    Dim a As Range
    Set a = onePager_Range_RBPC
    
    
    ' sprawdzimy po kolei czy nie ma przypadkiem juz takiego wpisu
    ' elementow: 4
    
    Dim czyZnalazlesDuplikat As Boolean
    czyZnalazlesDuplikat = False
    
    For x = 1 To 4
        czyZnalazlesDuplikat = innerCheckForRBPCH_LINE(1, rf, a)
    Next x
    
    checkIfNoDuplicates = Not czyZnalazlesDuplikat
End Function

Private Function innerCheckForRBPCH_LINE(line, rf As Range, a As Range)
    innerCheckForRBPCH_LINE = False
    
    
    ' offset cfg
    ' 0 - build phase
    ' 1 - # of veh
    ' 2 - tbw
    ' 3 - order release ch
    ' 4 - comments
    
    
    rfTempTxt = "" & _
        CStr(rf.Offset(0, SIXP.e_recent_bp_ch_comment - 1)) & _
        CStr(rf.Offset(0, SIXP.e_recent_bp_ch_no_of_veh - 1)) & _
        CStr(rf.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1)) & _
        CStr(rf.Offset(0, SIXP.e_recent_bp_ch_tbw - 1))
        
    aTempTxt = "" & _
        CStr(a.Offset(Int(line) - 1, 4)) & _
        CStr(a.Offset(Int(line) - 1, 1)) & _
        CStr(a.Offset(Int(line) - 1, 3)) & _
        CStr(a.Offset(Int(line) - 1, 2))
        
    If CStr(aTempTxt) = CStr(rfTempTxt) Then
        innerCheckForRBPCH_LINE = True
    End If
    
    
End Function

Private Sub wypelnienie_pcs_totals(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim third_chart(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    ' 1.
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna proj
        If Trim(rf) = Trim(tl.project) Then
            ' 2 kolumna plt order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
                ' 3 kolumna faza
                If Trim(rf.Offset(0, 2)) = Trim(tl.faza) Then
            
                    ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                    ' o ktory sami poprosilismy :)
                    ' dzieki temu ze wykorzystujemy zapis yyyycw
                    ' mozemy korzystac ze zwyklej arytmetyki >< :)
                    If CLng(rf.Offset(0, 3)) <= CLng(tl.cw) Then
                        Set third_chart(xx) = rf
                        xx = xx + 1
                    End If
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    ' 2.
    bubble_sort third_chart, 3, E_DESCENDING
    
    
    ' 3.
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    '''
    ''
    '
    
    i = 0
    ' poczatek danych order release na one pagerze
    ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm).Range("C6:N10").Clear
    'Dim pd As Range
    Set pD = ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm).Range("C6")

    For xxx = LBound(third_chart) To UBound(third_chart)
    
    
        Set rf = third_chart(xxx)
        
        If Not rf Is Nothing Then
        
            ' Debug.Print rf.Offset(0, 3)
        
        
            If pD.Value = "" Then
            
                pD.Value = rf.Offset(0, SIXP.e_5p_total - 1)
                ' N/A
                pD.Offset(0, 1).Value = rf.Offset(0, SIXP.e_5p_na - 1)
                
                ' ITDC
                pD.Offset(0, 2).Value = rf.Offset(0, SIXP.e_5p_itdc - 1)
                
                ' PNOC
                pD.Offset(0, 3).Value = rf.Offset(0, SIXP.e_5p_pnoc - 1)
                
                ' FMA OSEA
                pD.Offset(0, 4).Value = rf.Offset(0, SIXP.e_5p_fma_osea - 1)
                ' FMA EUR
                pD.Offset(1, 4).Value = rf.Offset(0, SIXP.e_5p_fma_eur - 1)
                
                ' ORDERED
                pD.Offset(0, 6).Value = rf.Offset(0, SIXP.e_5p_ordered - 1)
                
                ' TRANSIT arrived / in transit / future
                pD.Offset(0, 8).Value = rf.Offset(0, SIXP.e_5p_arrived - 1)
                pD.Offset(1, 8).Value = rf.Offset(0, SIXP.e_5p_in_transit - 1)
                pD.Offset(2, 8).Value = rf.Offset(0, SIXP.e_5p_future - 1)
                
                
                ' ppaps
                pD.Offset(0, 10).Value = rf.Offset(0, SIXP.e_5p_ppap_status - 1)
                pD.Offset(0, 11).Value = rf.Offset(0, SIXP.e_5p_no_ppap_status - 1)
                    
            ElseIf pD.Value <> "" Then
                Exit For
            End If
    
            
        Else
            Exit For
        End If
            
    Next xxx
    
    '
    ''
    '''
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    
    ' ------------------------------------------------------------------------------------------------
End Sub


Private Sub wypelnienie_pcs_xq(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    


    Dim flagFound As Boolean
    flagFound = False
    
    Dim ppapGateDate(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    
    ' 1.
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna proj
        If Trim(rf) = Trim(tl.project) Then
            ' 2 kolumna plt order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
                ' 3 kolumna faza
                If Trim(rf.Offset(0, 2)) = Trim(tl.faza) Then
            
                    If CLng(rf.Offset(0, 3)) = CLng(tl.cw) Then
                    
                        If Trim(rf) <> "" Then
                            Set ppapGateDate(xx) = rf
                            flagFound = True
                            
                            xx = xx + 1
                        End If
                    End If
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    
    ' 2.
    bubble_sort ppapGateDate, 3, E_DESCENDING

    
    Set pD = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("G23")
    pD.Value = "#" ' tylko w tym przypadku tak to wyglda
    
    'If flagFound Then pD.Value = CStr(ppapGateDate.Value)
    'If Not flagFound Then pD.Value = "#"
    
    For xxx = LBound(ppapGateDate) To UBound(ppapGateDate)
    
    
        Set rf = ppapGateDate(xxx)
        
        If Not rf Is Nothing Then
        
            ' Debug.Print rf.Offset(0, 3)
        
        
            If pD.Value = "#" Then
            
                pD.Value = CDate(rf.Offset(0, SIXP.e_xq_ppap_gate - 1))
                
                    
            ElseIf pD.Value <> "" Then
                Exit For
            End If
    
            
        Else
            Exit For
        End If
            
    Next xxx
    
End Sub



Private Sub wypelnienie_pcs_del_conf(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    'Dim pd As Range, sh As Worksheet
    
    Dim ssh As Worksheet
    Set ssh = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
    
    
    With ssh
        ' zera na del confie
        ' ------------------
        
        ' clear all
        .Range("K16:M28").Value = vbNullString
        
    End With
    
    
    ' GREEN GROUP
    ' ----------------------------------------------------------------
    ' T25 - on stock
    Set pD = ssh.Range("gonstock")
    pD.Value = assign_qty_to_on_stock(tl)
    
    ' T26 - EDI
    Set pD = ssh.Range("gedi")
    pD.Value = assign_qty_to_edi(tl)
    
    ' T27 - H/O
    Set pD = ssh.Range("gho")
    pD.Value = assign_qty_to_h_slash_o(tl)
    
    ' T28 - N/A
    Set pD = ssh.Range("gna")
    pD.Value = assign_qty_to_n_slash_a(tl)
    ' ----------------------------------------------------------------
    
    ' MRD FOR / AFTER GROUP
    ' ----------------------------------------------------------------
    
    ' AA25 - MRD
    Set pD = ssh.Range("gmrd")
    pD.Value = assign_qty_to_mrd_for_green(tl)
    ssh.Range("ymrd").Value = assign_qty_to_mrd_after_yellow(tl)
    ssh.Range("rmrd").Value = assign_qty_to_mrd_after_red(tl)
    
    ' AA26 - MRD alt
    Set pD = ssh.Range("galt")
    pD.Value = assign_qty_to_mrd_alt_for_green(tl)
    ssh.Range("yalt").Value = assign_qty_to_mrd_alt_after_yellow(tl)
    ssh.Range("ralt").Value = assign_qty_to_mrd_alt_after_red(tl)
    
    
    ' AA27 - MRD stagg
    Set pD = ssh.Range("gsmrd")
    pD.Value = assign_qty_to_mrd_staggered_for_green(tl)
    ssh.Range("ysmrd").Value = assign_qty_to_mrd_staggered_after_yellow(tl)
    ssh.Range("rsmrd").Value = assign_qty_to_mrd_staggered_after_red(tl)
    
    
    ' AA28 - MRD stagg oncost
    Set pD = ssh.Range("gsoncost")
    pD.Value = assign_qty_to_mrd_staggered_oncost_for_green(tl)
    ssh.Range("ysoncost").Value = assign_qty_to_mrd_staggered_oncost_after_yellow(tl)
    ssh.Range("rsoncost").Value = assign_qty_to_mrd_staggered_oncost_after_red(tl)
    
    ' AA29 - MRD oncost
    Set pD = ssh.Range("goncost")
    pD.Value = assign_qty_to_mrd_oncost_for_green(tl)
    ssh.Range("yoncost").Value = assign_qty_to_mrd_oncost_after_yellow(tl)
    ssh.Range("roncost").Value = assign_qty_to_mrd_oncost_after_red(tl)
    
    
    
    ' alt two
    Set pD = ssh.Range("galttwo")
    pD.Value = assign_qty_to_mrd_alt_two_for_green(tl)
    ssh.Range("yalttwo").Value = assign_qty_to_mrd_alt_two_after_yellow(tl)
    ssh.Range("ralttwo").Value = assign_qty_to_mrd_alt_two_after_red(tl)
    
    
    
    Set pD = ssh.Range("gsalttwo")
    pD.Value = assign_qty_to_mrd_staggered_alt_two_for_green(tl)
    ssh.Range("ysalttwo").Value = assign_qty_to_mrd_staggered_alt_two_after_yellow(tl)
    ssh.Range("rsalttwo").Value = assign_qty_to_mrd_staggered_alt_two_after_red(tl)
    
    ' ----------------------------------------------------------------
    
    
    ' RED GROUP
    ' ----------------------------------------------------------------
    
    ' T30 - open - to jest zle!
    'Set pD = ssh.Range("yopen")
    'pD.Value = assign_qty_to_open(tl)
    
    ' T31 - pot itdc
    Set pD = ssh.Range("rpotitdc")
    pD.Value = assign_qty_to_pot_itdc(tl)
    
    
    
    Set pD = ssh.Range("yopen")
    pD.Value = assign_qty_to_yellow_open(tl)
    
    Set pD = ssh.Range("ropen")
    pD.Value = assign_qty_to_red_open(tl)
    
    
    
    ' ----------------------------------------------------------------
    
    ' ------------------------------------------------------------------------------------------------
    
    
    ' now clear zeros
    Dim delConfGroup As Range
    Set delConfGroup = ssh.Range("K16:M28")
    
    For Each ir In delConfGroup
        If IsNumeric(ir.Value) Then
            If Int(ir.Value) = 0 Then
                ir.Value = vbNullString
            End If
        End If
    Next ir
End Sub

Private Sub wypelnienie_pcs_resp(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    Dim ssh As Worksheet
    Set ssh = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
    
    
    ' resp
    ' ----------------------------------------------------------------
    
    ' PEM
    Set pD = ssh.Range("resppem")
    ' pD.Value = "      " & find_PEM_for_this_rep(tl)
    pD.Value = "" & find_PEM_for_this_rep(tl)
    
    ' PPM
    Set pD = ssh.Range("respppm")
    pD.Value = "" & find_PPM_for_this_rep(tl)
    
    ' FMA
    Set pD = ssh.Range("respfma")
    pD.Value = "" & find_FMA_for_this_rep(tl)
    
    ' SQE
    Set pD = ssh.Range("respsqe")
    pD.Value = "" & find_SQE_for_this_rep(tl)
    
    ' SQE
    Set pD = ssh.Range("resposea")
    pD.Value = "" & find_OSEA_for_this_rep(tl)
    
    
    ' ----------------------------------------------------------------
    
    ' ------------------------------------------------------------------------------------------------
End Sub

Private Function find_PEM_for_this_rep(mtl As T_Link) As String
    find_PEM_for_this_rep = common_find_resp(e_resp_pem, mtl)
End Function

Private Function find_PPM_for_this_rep(mtl As T_Link) As String
    find_PPM_for_this_rep = common_find_resp(e_resp_ppm, mtl)
End Function

Private Function find_FMA_for_this_rep(mtl As T_Link) As String
    find_FMA_for_this_rep = common_find_resp(e_resp_fma, mtl)
End Function

Private Function find_SQE_for_this_rep(mtl As T_Link) As String
    find_SQE_for_this_rep = common_find_resp(e_resp_sqe, mtl)
End Function

Private Function find_OSEA_for_this_rep(mtl As T_Link) As String
    find_OSEA_for_this_rep = common_find_resp(e_resp_osea, mtl)
End Function


Private Function common_find_resp(resp As E_RESP_ORDER, mtl As T_Link) As String
    
    common_find_resp = ""
    
    Dim sh As Worksheet, r As Range
    Set sh = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
    
    
    If sh.Range("A2") <> "" Then
        Set r = sh.Range("A1").End(xlDown)
    Else
        Set r = sh.Range("A2")
    End If
    
    Do
    
        
        
        
        If Trim(r) = Trim(mtl.project) Then
            If Trim(r.Offset(0, 1)) = Trim(mtl.plt) Then
                If Trim(r.Offset(0, 2)) = Trim(mtl.faza) Then
                    If CLng(r.Offset(0, 3)) <= CLng(mtl.cw) Then
                        common_find_resp = r.Offset(0, resp - 1)
                        
                        If Trim(common_find_resp) <> "" Then
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    
    
        Set r = r.Offset(-1, 0)
    Loop Until r.Row = 1
End Function

Private Sub wypelnienie_pcs_open_issues(rf As Range, pD As Range, sh As Worksheet, tl As T_Link)

    ' ------------------------------------------------------------------------------------------------
    
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim open_issues_arr(0 To 512) As Range
    
    Set rf = sh.Cells(2, 1)
    
    
    ' 1.
    xx = 0
    Do
        If Trim(rf) = "" Then
            Exit Do
        End If
        
        
        ' 1 kolumna proj
        If Trim(rf) = Trim(tl.project) Then
            ' 2 kolumna plt order release status
            If Trim(rf.Offset(0, 1)) = Trim(tl.plt) Then
            
                ' 3 kolumna faza
                If Trim(rf.Offset(0, 2)) = Trim(tl.faza) Then
            
                    ' plus patrzymymy sobie na wszystkie update'y do ostatniego cw
                    ' o ktory sami poprosilismy :)
                    ' dzieki temu ze wykorzystujemy zapis yyyycw
                    ' mozemy korzystac ze zwyklej arytmetyki >< :)
                    If CLng(rf.Offset(0, 3)) <= CLng(tl.cw) Then
                    
                        If CLng(rf.Parent.Cells(rf.Row, SIXP.e_open_issues_visible).Value) = 1 Then
                            Set open_issues_arr(xx) = rf
                            xx = xx + 1
                        End If
                    End If
                End If
            End If
            
        End If
        
        Set rf = rf.Offset(1, 0)
    Loop Until Trim(rf) = ""
    
    ' 2.
    bubble_sort open_issues_arr, 3, E_DESCENDING
    
    
    ' 3.
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    '''
    ''
    '
    
    i = 0
    ' poczatek danych order release na one pagerze
    ' Dim pd As Range
    
    ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("P5:Y419").Value = ""
    
    Set pD = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("P5")
    ' recent_bp_changes()
    For xxx = LBound(open_issues_arr) To UBound(open_issues_arr)
    
    
        Set rf = open_issues_arr(xxx)
        
        If Not rf Is Nothing Then
    
            ' warunek czy juz nie przekorczylismy wielkosci tabelki
            ' dzielenie przez dwa poniewaz sa mergi i w sumie przeskakuje na co drugi wiersz
            ' sumarycznie contentu jest 4 jednak wizualnie zaafektowane jest 8 wierszy arkusza
            If Int(i) < SIXP.G_OPEN_ISSUES_LIMIT Then
                 
                ' STATUS 2 kolumny
                ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("P" & CStr(5 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_status - 1)
                
                ' # of pn 2 kolumny
                ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("R" & CStr(5 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_no_of_pn - 1)
                
                ' part  / supplier 3 kolumny
                ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("T" & CStr(5 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_part_supplier - 1)
                
                ' delivery 2 kolumny
                ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("W" & CStr(5 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_delivery - 1)
                
                ' cmnt 15 kolumn
                ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("Y" & CStr(5 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_comment - 1)
                
                i = i + 2
    
            Else
                Exit For
            End If
        Else
            Exit For
        End If
            
    Next xxx
    
    '
    ''
    '''
    ' ======================================================================================================
    ' ======================================================================================================
    ' ======================================================================================================
    
    ' ------------------------------------------------------------------------------------------------
End Sub


