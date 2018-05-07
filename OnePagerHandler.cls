VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OnePagerHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
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



Public Function get_one_pager_sheet() As Worksheet
    Set get_one_pager_sheet = one_pager_sheet
End Function


Private Sub Class_Initialize()
    Set kolekcja_linkow = New Collection
    
    
    ' procedura z moduly (public sub!)
    clear_one_pager
    
    clear_new_one_pager
    
    
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
    
    ' 10 == xq
    Set pcs(10) = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
    
    
    Set one_pager_sheet = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm)
    
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

    Dim Sh As Worksheet, r As Range, rf As Range, pD As Range
    
    inicjuj_power_pointa_jesli_taka_jest_konfiguracja e

    Dim iter As Long
    iter = 1
    For Each link_raportu In kolekcja_linkow
    
        
        
        Set r = Nothing
        
        ' cialo generowania iterowanego one pagera
        ' =============================================
        For x = LBound(pcs) To UBound(pcs)
            Set Sh = pcs(x)
            Sh.Activate
            SIXP.LoadingFormModule.incLoadingForm
            
            ' to jest najprostsza logika dzialania jednak to nie jest tak chocby
            ' dla pcs(2) to jest dla danych order release status trzeba sprawdzic nawet wszystkie fazy dla wybranego projektu
            ' wiec odrobine smutek - gdyz trzeba rozszerzyc dzialanie mocno...
            
            
            
            ' patrzymy tylko w main == x == 1
            If x = 1 Then Set r = link_raportu.znajdz_siebie_w_arkuszu(Sh)
            
            If Not r Is Nothing Then
                ' r.Select
                ' MsgBox sh.Name & " " & r & " " & r.Offset(0, 1) & " " & r.Offset(0, 3)
                
                ' niech narazie w ramach wersji testowej robia sie aktywacje oraz selekcje
                ' --------------------------------------------------------------------------
                ' --------------------------------------------------------------------------
                
                If x = SIXP.e_pcs_order_release_status Then
                    wypelnienie_pcs_order_release_status rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_pnoc Then
                    ' wypelnienie_pcs_pnoc rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_osea Then
                    wypelnienie_pcs_osea rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_recent_build_plan_changes Then
                    wypelnienie_pcs_recent_build_plan_changes rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_totals Then
                    wypelnienie_pcs_totals rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_delivery_confirmation Then
                    wypelnienie_pcs_del_conf rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_resp Then
                    wypelnienie_pcs_resp rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_open_issues Then
                    wypelnienie_pcs_open_issues rf, pD, Sh, link_raportu
                ElseIf x = SIXP.e_pcs_xq Then
                    wypelnienie_pcs_xq rf, pD, Sh, link_raportu
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
End Sub

Private Sub szary_na_tym_raporcie(l As T_Link)
    
    Dim orsRange As Range
    Set orsRange = ThisWorkbook.Sheets("ONE PAGER").Range("A3:A10")
    
    Dim r As Range
    
    
    For Each r In orsRange
    
        If CStr(UCase(r.Value)) = CStr(UCase(l.faza)) Then
            r.Parent.Range(r, r.Offset(0, 14)).Interior.Color = RGB(200, 200, 200)
        Else
            r.Parent.Range(r, r.Offset(0, 14)).Interior.Color = RGB(255, 255, 255)
        End If
    Next r
    
End Sub

Private Sub przesortuj_raz_jeszcze_order_release_status_w_zgodzie_z_phase_list(l As T_Link)
    
    Dim orsRange As Range
    Set orsRange = ThisWorkbook.Sheets("ONE PAGER").Range("A3:A10")
    
    
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
                
                    For x = 1 To 14
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





Private Function assign_qty_to_mrd_oncost_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_oncost_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_oncostmrd))
End Function

Private Function assign_qty_to_mrd_oncost_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_oncost_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_oncostmrd))
End Function


Private Function assign_qty_to_mrd_alt_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_alttwomrd))
End Function

Private Function assign_qty_to_mrd_alt_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_alt_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_alttwomrd))
End Function




Private Function assign_qty_to_mrd_staggered_alt_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_alt_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_salttwomrd))
End Function

Private Function assign_qty_to_mrd_staggered_alt_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_staggered_alt_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_salttwomrd))
End Function





Private Function assign_qty_to_mrd_two_for_green(mtl As T_Link) As Long
    assign_qty_to_mrd_two_for_green = CLng(common_assign_del_conf_qty(mtl, e_del_conf_for_twomrd))
End Function

Private Function assign_qty_to_mrd_two_after_yellow(mtl As T_Link) As Long
    assign_qty_to_mrd_two_after_yellow = CLng(common_assign_del_conf_qty(mtl, e_del_conf_after_twomrd))
End Function



Private Function assign_qty_to_open(mtl As T_Link) As Long
    ' assign_qty_to_open = CLng(common_assign_del_conf_qty(mtl, e_del_conf_open))
    
    
    tmp = 0
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_alt))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_alttwomrd))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_mrd))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_oncostmrd))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_salttwomrd))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_smrd))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_red_after_soncostmrd))
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_yellow_open))
    
    tmp = tmp + CLng(common_assign_del_conf_qty(mtl, e_del_conf_open))
    
    
    
    assign_qty_to_open = tmp
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
    If e = E_ONE_PAGERS_INTO_POWER_POINT Then
    
        
        myOwnIter = iter * 2 - 1
    
        DoEvents
        makeSlide CLng(myOwnIter), ppApp, pp, myCustomLayout, l

    ElseIf e = E_ONE_PAGERS_INTO_SEPERATE_EXCELS Then
    
        Dim w As Workbook, source_sh As Worksheet
        
        ' nowy pliczor
        Set w = Workbooks.Add
        
        ' tutaj akursz one pagera
        Set source_sh = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm)
        source_sh.Copy Before:=w.Sheets(1)
        
        ' pasuje dodac jeszcze arkusze zawierajace zrodla dla
        ' charty 1, 2, 3
        ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm).Copy Before:=w.Sheets(1)
        ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm).Copy Before:=w.Sheets(1)
        ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm).Copy Before:=w.Sheets(1)
        
        
        Dim obiekty As ChartObjects
        Dim obiekt_wykresu As ChartObject, source_range As Range, rep_sh As Worksheet
        Dim wykres As Excel.Chart
        
        
        Set rep_sh = w.Sheets(SIXP.G_one_pager_sh_nm)
        rep_sh.Activate
        
        Set obiekty = rep_sh.ChartObjects
        
        For Each obiekt_wykresu In obiekty
        
        
            'Set obiekt_wykresu = Nothing
            obiekt_wykresu.Activate
            Set wykres = Nothing
            On Error Resume Next
            Set wykres = obiekt_wykresu.Chart
            
            
            If Not wykres Is Nothing Then
            
                If CStr(wykres.name) Like "*" & CStr(SIXP.G_CHART_1_NM) & "*" Then
                
                
                    
                    ' =CHART1_Handler!$C$5:$E$12
                    wykres.SetSourceData Source:=w.Sheets(SIXP.G_chart1_handler_sh_nm).Range("C5:E12")
                    
                ElseIf CStr(wykres.name) Like "*" & CStr(SIXP.G_CHART_2_NM) & "*" Then
                
                    ' =CHART2_Handler!$C$5:$I$6
                    wykres.SetSourceData Source:=w.Sheets(SIXP.G_chart2_handler_sh_nm).Range("C5:I16")
                
                ElseIf CStr(wykres.name) Like "*" & CStr(SIXP.G_CHART_3_NM) & "*" Then
                
                    ' =CHART3_Handler!$C$5:$N$8
                    wykres.SetSourceData Source:=w.Sheets(SIXP.G_chart3_handler_sh_nm).Range("C5:N8")
                
                
                Else
                
                    ' no operation
                    
                End If
            End If
        Next obiekt_wykresu
        
        
        
        ' plik w to jest nowy workbook w ktorym znajduje sie pojedynczy raport dla jednej "linijki" z 6P
        
        Application.DisplayAlerts = False
        w.SaveAs ThisWorkbook.Path & Application.PathSeparator & _
            CStr(SIXP.G_EXCEL_REPORT_PREFIX & CStr(l.plt) & "_" & CStr(l.project) & "_" & CStr(l.faza) & "_" & CStr(l.cw)), _
            51, , , , , , xlLocalSessionChanges
        Application.DisplayAlerts = True
        
        
        w.Sheets(SIXP.G_one_pager_sh_nm).Activate
        
        
        ' to gdy pracujemy na outlooku
        '.Attachments.Add w.FullName
        ' nie chcemy zamykac niech sobie wisi
        ' w.Close False
    Else
        MsgBox "ta sytuacja nie jest mozliwa - blad krytyczny - makro zatrzymalo sie!"
        End
    End If
End Sub


Private Sub makeSlide(ktorySlide As Long, ppApp As PowerPoint.Application, _
    pp As PowerPoint.Presentation, myCustomLayout As PowerPoint.CustomLayout, _
    Optional l As T_Link)

    

    ' pp.PageSetup.SlideSize = ppSlideSizeA4Paper
    ' pp.PageSetup.SlideOrientation = msoOrientationVertical
    pp.PageSetup.SlideHeight = 297 * 2.88
    pp.PageSetup.slideWidth = 210 * 2.88
    copyPasteDataFromOnePager 1, CLng(ktorySlide), ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("A1:AC42"), l
    copyPasteDataFromOnePager 2, CLng(ktorySlide + 1), ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AF1:BH42"), l

    
End Sub


Private Sub copyPasteDataFromOnePager(page As Integer, _
    slideCounter As Long, _
    copiedRange As Range, _
    Optional l As T_Link)
    
    Dim toCopy As Range, shToCopy As Worksheet
    Set ppSlide = pp.Slides.AddSlide(slideCounter, myCustomLayout)
    
    Set shToCopy = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm)
    shToCopy.DisplayPageBreaks = False
    
    ' copiedRange
    Set toCopy = copiedRange
    
    toCopy.CopyPicture
    ppSlide.Shapes.Paste
    
    Dim myShape As PowerPoint.Shape
    Set myShape = ppSlide.Shapes(1)
        
    Do
        With myShape
        
            .LockAspectRatio = msoTrue
            
            ' in march 2018
            ' G_OFFSET_FROM_TOP_IN_POWERPOINT_SLIDE: 50
            
            If .Height > (pp.PageSetup.SlideHeight - SIXP.G_OFFSET_FROM_TOP_IN_POWERPOINT_SLIDE) Then
                .Height = pp.PageSetup.SlideHeight - SIXP.G_OFFSET_FROM_TOP_IN_POWERPOINT_SLIDE
            ElseIf .Width > pp.PageSetup.slideWidth Then
                .Width = pp.PageSetup.slideWidth - SIXP.G_OFFSET_IN_POWERPOINT_SLIDE - SIXP.G_OFFSET_IN_POWERPOINT_SLIDE
            End If
            
            .LockAspectRatio = msoTrue
            .Left = przelicz(pp.PageSetup.slideWidth, .Width) + SIXP.G_OFFSET_IN_POWERPOINT_SLIDE
            .Top = SIXP.G_OFFSET_FROM_TOP_IN_POWERPOINT_SLIDE
        End With
        
        If Not ((myShape.Height > pp.PageSetup.SlideHeight) Or (myShape.Width > pp.PageSetup.slideWidth)) Then
            Exit Do
        End If
    Loop While True
    
    
    ' l as T_Link
    
    Dim lr As Linker
    Set lr = New Linker
    
    If Not l Is Nothing Then
        ppSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 10, 10, _
            pp.PageSetup.slideWidth - SIXP.G_OFFSET_IN_POWERPOINT_SLIDE, 30).TextFrame.TextRange.Text = _
                CStr(lr.return_full_concated_r_string_comma_seperated(l))
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

Private Sub wypelnienie_pcs_order_release_status(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
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
    
    Set rf = Sh.Cells(2, 1)
    
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
    
    
    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("A3:O10").Value = ""
    
    Set pD = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("A3")
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
                If warunek_czy_jest_dana_faza_sie_nie_pojawila(ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("A3:A10"), rf.Offset(0, SIXP.e_link_faza - 1)) Then
                    
                    ' PHASE
                    pD.Offset(i, 0) = rf.Offset(0, SIXP.e_link_faza - 1)
                    
                    ' MRD
                    pD.Offset(i, 1) = rf.Offset(0, SIXP.e_order_release_mrd - 1)
                    
                    ' Build
                    pD.Offset(i, 3) = rf.Offset(0, SIXP.e_order_release_build - 1)
                    
                    ' bom freeze
                    pD.Offset(i, 5) = rf.Offset(0, SIXP.e_order_release_bom_freeze - 1)
                    
                    ' # of veh
                    pD.Offset(i, 7) = rf.Offset(0, SIXP.e_order_release_no_of_veh - 1)
                    
                    ' orders due
                    pD.Offset(i, 9) = rf.Offset(0, SIXP.e_order_release_orders_due - 1)
                    
                    ' released
                    pD.Offset(i, 11) = rf.Offset(0, SIXP.e_order_release_released - 1)
                    
                    ' weeks delay
                    pD.Offset(i, 13) = rf.Offset(0, SIXP.e_order_release_weeks_delay - 1)
                    
                    
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

Private Sub wypelnienie_pcs_pnoc(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' te dane sa pod wykres wraz z danymi historycznymi
    ' - zatem patrzymy na ten sam projekt z perspektywy
    ' roznych CW
    ' ------------------------------------------------------------------------------------------------
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim fst_chart(0 To 512) As Range
    
    Set rf = Sh.Cells(2, 1)
    
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

Private Sub wypelnienie_pcs_osea(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim sec_chart(0 To 512) As Range
    
    Set rf = Sh.Cells(2, 1)
    
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

Private Sub wypelnienie_pcs_recent_build_plan_changes(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim recent_bp_changes(0 To 512) As Range
    
    Set rf = Sh.Cells(2, 1)
    
    
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
    
    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("R3:Z9").Value = ""
    
    Set pD = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("R3")
    ' recent_bp_changes()
    For xxx = LBound(recent_bp_changes) To UBound(recent_bp_changes)
    
    
        Set rf = recent_bp_changes(xxx)
        
        If Not rf Is Nothing Then
    
            ' warunek czy juz nie przekorczylismy wielkosci tabelki
            ' dzielenie przez dwa poniewaz sa mergi i w sumie przeskakuje na co drugi wiersz
            ' sumarycznie contentu jest 4 jednak wizualnie zaafektowane jest 8 wierszy arkusza
            If Int(i / 2#) < SIXP.G_RECENT_BUILD_PLAN_CHANGES_LIMIT Then
            
            
                If isThereAnyContent(rf) Then
                     
                    ' PHASE
                    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("R" & CStr(3 + i)) = _
                        rf.Offset(0, SIXP.e_link_faza - 1)
                    
                    ' # of veh
                    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("T" & CStr(3 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_no_of_veh - 1)
                    
                    ' TBW
                    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("V" & CStr(3 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_tbw - 1)
                    
                    ' order release change
                    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("X" & CStr(3 + i)) = _
                        rf.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1)
                    
                    ' cmnt
                    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("Z" & CStr(3 + i)) = _
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

Private Sub wypelnienie_pcs_xq(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    


    Dim flagFound As Boolean
    flagFound = False
    
    Dim ppapGateDate(0 To 512) As Range
    
    Set rf = Sh.Cells(2, 1)
    
    
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

    
    Set pD = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("G23")
    pD.Value = "#" ' tylko w tym przypadku tak to wyglda
    
    'If flagFound Then pD.Value = CStr(ppapGateDate.Value)
    'If Not flagFound Then pD.Value = "#"
    
    For xxx = LBound(ppapGateDate) To UBound(ppapGateDate)
    
    
        Set rf = ppapGateDate(xxx)
        
        If Not rf Is Nothing Then
        
            ' Debug.Print rf.Offset(0, 3)
        
        
            If pD.Value = "#" Then
            
                pD.Value = rf.Offset(0, SIXP.e_xq_ppap_gate - 1)
                
                    
            ElseIf pD.Value <> "" Then
                Exit For
            End If
    
            
        Else
            Exit For
        End If
            
    Next xxx
    
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

Private Sub wypelnienie_pcs_totals(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim third_chart(0 To 512) As Range
    
    Set rf = Sh.Cells(2, 1)
    
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

Private Sub wypelnienie_pcs_del_conf(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    'Dim pd As Range, sh As Worksheet
    
    Dim ssh As Worksheet
    Set ssh = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm)
    
    
    With ssh
        ' zera na del confie
        ' ------------------
        
        ' calci after & before
        .Range("AA25:AA31") = 0
        .Range("AC25:AC31") = 0
        
        
        ' green
        .Range("T25:T28") = 0
        
        
        ' red
        .Range("T30:T32") = 0
    End With
    
    
    ' GREEN GROUP
    ' ----------------------------------------------------------------
    ' T25 - on stock
    Set pD = ssh.Range("T25")
    pD.Value = assign_qty_to_on_stock(tl)
    
    ' T26 - EDI
    Set pD = ssh.Range("T26")
    pD.Value = assign_qty_to_edi(tl)
    
    ' T27 - H/O
    Set pD = ssh.Range("T27")
    pD.Value = assign_qty_to_h_slash_o(tl)
    
    ' T28 - N/A
    Set pD = ssh.Range("T28")
    pD.Value = assign_qty_to_n_slash_a(tl)
    ' ----------------------------------------------------------------
    
    ' MRD FOR / AFTER GROUP
    ' ----------------------------------------------------------------
    
    ' AA25 - MRD
    Set pD = ssh.Range("AA25")
    pD.Value = assign_qty_to_mrd_for_green(tl)
    ssh.Range("AC25").Value = assign_qty_to_mrd_after_yellow(tl)
    
    ' AA26 - MRD alt
    Set pD = ssh.Range("AA26")
    pD.Value = assign_qty_to_mrd_alt_for_green(tl)
    ssh.Range("AC26").Value = assign_qty_to_mrd_alt_after_yellow(tl)
    
    
    ' AA27 - MRD stagg
    Set pD = ssh.Range("AA27")
    pD.Value = assign_qty_to_mrd_staggered_for_green(tl)
    ssh.Range("AC27").Value = assign_qty_to_mrd_staggered_after_yellow(tl)
    
    ' AA28 - MRD stagg oncost
    Set pD = ssh.Range("AA28")
    pD.Value = assign_qty_to_mrd_staggered_oncost_for_green(tl)
    ssh.Range("AC28").Value = assign_qty_to_mrd_staggered_oncost_after_yellow(tl)
    
    ' AA29 - MRD oncost
    Set pD = ssh.Range("AA29")
    pD.Value = assign_qty_to_mrd_oncost_for_green(tl)
    ssh.Range("AC29").Value = assign_qty_to_mrd_oncost_after_yellow(tl)
    
    ' AA30 - MRD oncost
    Set pD = ssh.Range("AA30")
    pD.Value = assign_qty_to_mrd_alt_two_for_green(tl)
    ssh.Range("AC30").Value = assign_qty_to_mrd_alt_two_after_yellow(tl)
    
    
    ' AA31 - MRD oncost
    Set pD = ssh.Range("AA31")
    pD.Value = assign_qty_to_mrd_staggered_alt_two_for_green(tl)
    ssh.Range("AC31").Value = assign_qty_to_mrd_staggered_alt_two_after_yellow(tl)
    
    ' AA28 - MRD stagg two - obsolete
    'Set pD = ssh.Range("AA28")
    'pD.Value = assign_qty_to_mrd_staggered_two_for_green(tl)
    'ssh.Range("AC28").Value = assign_qty_to_mrd_staggered_two_after_yellow(tl)
    
    ' AA29 - MRD two - obsolete
    'Set pD = ssh.Range("AA29")
    'pD.Value = assign_qty_to_mrd_two_for_green(tl)
    'ssh.Range("AC29").Value = assign_qty_to_mrd_two_after_yellow(tl)
    ' ----------------------------------------------------------------
    
    
    ' RED GROUP
    ' ----------------------------------------------------------------
    
    ' T30 - open
    Set pD = ssh.Range("T30")
    pD.Value = assign_qty_to_open(tl)
    
    ' T31 - pot itdc
    Set pD = ssh.Range("T31")
    pD.Value = assign_qty_to_pot_itdc(tl)
    
    ' T32 - undef / too late
    Set pD = ssh.Range("T32")
    ' too late
    pD.Value = assign_qty_to_too_late(tl)
    
    ' ----------------------------------------------------------------
    
    ' ------------------------------------------------------------------------------------------------
End Sub

Private Sub wypelnienie_pcs_resp(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)
    ' ------------------------------------------------------------------------------------------------
    
    Dim ssh As Worksheet
    Set ssh = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm)
    
    
    ' resp
    ' ----------------------------------------------------------------
    ' a37 PEM
    Set pD = ssh.Range("A37")
    pD.Value = "PEM: " & find_PEM_for_this_rep(tl)
    
    ' k37 PPM
    Set pD = ssh.Range("K37")
    pD.Value = "PPM: " & find_PPM_for_this_rep(tl)
    
    ' u37 FMA
    Set pD = ssh.Range("U37")
    pD.Value = "FMA: " & find_FMA_for_this_rep(tl)
    
    'pd.Select
    
    ' a39 SQE
    Set pD = ssh.Range("A39")
    pD.Value = "SQE: " & find_SQE_for_this_rep(tl)
    
    ' k39 SQE
    Set pD = ssh.Range("K39")
    pD.Value = "OSEA: " & find_OSEA_for_this_rep(tl)
    
    
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
    
    Dim Sh As Worksheet, r As Range
    Set Sh = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
    
    
    If Sh.Range("A2") <> "" Then
        Set r = Sh.Range("A1").End(xlDown)
    Else
        Set r = Sh.Range("A2")
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

Private Sub wypelnienie_pcs_open_issues(rf As Range, pD As Range, Sh As Worksheet, tl As T_Link)

    ' ------------------------------------------------------------------------------------------------
    
    
    ' 1. skompletuj dane,
    ' 2. posortuj dane,
    ' 3. wez najmlodsze itemy (najwieksze wartosciowo - daty jako liczby - im liczba wyzsza tym data mlodsza)
    
    Dim open_issues_arr(0 To 512) As Range
    
    Set rf = Sh.Cells(2, 1)
    
    
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
    
    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AG3:AP25").Value = ""
    
    Set pD = ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AG3")
    ' recent_bp_changes()
    For xxx = LBound(open_issues_arr) To UBound(open_issues_arr)
    
    
        Set rf = open_issues_arr(xxx)
        
        If Not rf Is Nothing Then
    
            ' warunek czy juz nie przekorczylismy wielkosci tabelki
            ' dzielenie przez dwa poniewaz sa mergi i w sumie przeskakuje na co drugi wiersz
            ' sumarycznie contentu jest 4 jednak wizualnie zaafektowane jest 8 wierszy arkusza
            If Int(i) < SIXP.G_OPEN_ISSUES_LIMIT Then
                 
                ' STATUS 2 kolumny
                ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AG" & CStr(3 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_status - 1)
                
                ' # of pn 2 kolumny
                ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AI" & CStr(3 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_no_of_pn - 1)
                
                ' part  / supplier 3 kolumny
                ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AK" & CStr(3 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_part_supplier - 1)
                
                ' delivery 2 kolumny
                ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AN" & CStr(3 + i)) = _
                    rf.Offset(0, SIXP.e_open_issues_delivery - 1)
                
                ' cmnt 15 kolumn
                ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Range("AP" & CStr(3 + i)) = _
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
