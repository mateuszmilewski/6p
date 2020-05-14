Attribute VB_Name = "ExternalSIXPTimeModule"
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



' to jest sub wykorzystywane do wygenerowania arkusza 6p time w wizardzie - my potrzebujemy bardziej mod
' tego rozwiazania do klasy
Public Sub inner_6p_time(mm, md, mp)
    
    ' sekcja bez pivotow
    ' dopasowanie do 6p
    ' ======================================================
    ''
    '
    
    SIXP.LoadingFormModule.showLoadingForm
    
    
    ' aby w ogole rozpoczac liczenie musze zrozumiec podstawowe definicje jakimi rzadzi sie poprzedni Quarter i w jakim cely
    ' mam w ogole zaciagac dane
    
    ' najpierw zrobmy nowy arkusz do ktorego tak jak w pierwszej generacji QT bedzie wsadzac kolejne dane
    ' jednak tym razem zrobimy to lepiej poniewaz z gory narzuce uklad kolumn taki jaki bedzie dostepny w nowym makrze 6p (nastepca Q)
    Dim w As Workbook, wrksh As Worksheet, m As Worksheet, d As Worksheet, p As Worksheet
    ' arkusz pusow bez dziur i na naszym arkuszu nowym z qt2
    ' Dim puses As Worksheet
    ' Set w = dodaj_nowy_arkusz()
    ' dodajemy nowy arkusz - nie wazne, czy sa tam jakies inne arkusze
    ' Set wrksh = wyodrebnij_arkusz_na_ktorym_bede_pracowal(w)
    Set wrksh = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    wrksh.Range("A1:ZZ100000").Clear
    ' Set puses = dodaj_nowy_arkusz_pusow(w)
    Set m = mm ' ThisWorkbook.Sheets(WizardMain.MASTER_SHEET_NAME)
    Set d = md ' ThisWorkbook.Sheets(WizardMain.DETAILS_SHEET_NAME)
    Set p = mp ' ThisWorkbook.Sheets(WizardMain.PICKUPS_SHEET_NAME)
    
    Dim m_r_d As Range, r_biw_ga As Range, build_start As Range
    Set m_r_d = d.Cells(SIXP.mrd, 2)
    Set build_start = d.Range("build_start") ' to samo powinno byc w D1 arkusza wiz buff
    
    ' also it will be in buff on Q1
    Set r_biw_ga = d.Cells(SIXP.biw_ga, 2)
    
    SIXP.LoadingFormModule.incLoadingForm
    
    
    fill_wiz_buff_with_all_details_but_transpose_from_o1 d, wrksh
    
    
    SIXP.LoadingFormModule.incLoadingForm

            
    
    
    'e_5p_total = 5
    'e_5p_na
    'e_5p_itdc
    'e_5p_pnoc
    'e_5p_fma_eur
    'e_5p_fma_osea
    'e_5p_ordered
    'e_5p_arrived
    'e_5p_in_transit
    'e_5p_future
    'e_5p_ppap_status
    'e_5p_no_ppap_status
    ' piaty pieces
    
    ' --------------------------------------------------------------------
    ' pierwsza linia od 3 kolumny mamy wolne miejsce
    
    ' Cells(3,1) jako MRD (CW)
    pierwszy_wiersz_pod_dane_ogolne = 1
    With wrksh
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 3)
            .Value = CStr(d.Cells(SIXP.mrd, 2))
            If .Comment Is Nothing Then
                .AddComment "MRD"
            End If
        End With
        
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 4)
            .Value = CStr(d.Cells(SIXP.build_start, 2))
            If .Comment Is Nothing Then
                .AddComment "BUILD START"
            End If
        End With
        
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 5)
            .Value = CStr(d.Cells(SIXP.build_end, 2))
            If .Comment Is Nothing Then
                .AddComment "BUILD END"
            End If
        End With
        
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 6)
            .Value = CStr(d.Cells(SIXP.BOM, 2))
            If .Comment Is Nothing Then
                .AddComment "BOM"
            End If
        End With
        
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 7)
            .Value = CStr(d.Cells(SIXP.ppap_gate, 2))
            If .Comment Is Nothing Then
                .AddComment "PPAP GATE"
            End If
        End With
        
        
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 9)
            .Value = CStr(Date)
            If .Comment Is Nothing Then
                .AddComment "Today"
            End If
        End With
        
        
        ' resp
        With .Cells(pierwszy_wiersz_pod_dane_ogolne, 10)
            .Value = CStr(d.Cells(SIXP.koordynator, 2))
            If .Comment Is Nothing Then
                .AddComment "FMA COORD"
            End If
        End With
        
    End With
    ' --------------------------------------------------------------------
    
    
    SIXP.LoadingFormModule.incLoadingForm
    
    wrksh.Cells(1, 1) = "6P"
    ' total
    wrksh.Cells(2, 1) = "TOTAL FMA*"
    
    ' pierwszy filtr odpwiedzialny jest za resp, drugi za kolumne przegladana :)
    wrksh.Cells(3, 1) = _
        iteruj_recur("*", 0, przelicz_zasieg(m, SIXP.pn, SIXP.Responsibility), "", E_NOT_EQUAL)
    
    
    
    SIXP.LoadingFormModule.incLoadingForm
    
    
    Dim rng As Range
    Set rng = wrksh.Cells(2, 2)
    
    rng.Offset(-1, 0) = "RESP"
    Set rng = zrob_recursy_dla("*", m, rng, SIXP.Responsibility)
    
    SIXP.RespAdjusterModule.resp_adjuster
    wrksh.Range("G1").Value = "IN SCOPE"
    wrksh.Range("H1").Value = CStr(podlicz_w_zgodzie_z_ukladem_z_arkusza_register())
    
    
    SIXP.LoadingFormModule.showLoadingForm
    SIXP.LoadingFormModule.incLoadingForm
    
    
    ' PRZYGOTOWANIE FILTROWANIA
    ' ========================================
    ' ========================================
    
    Dim fltr As String
    fltr = przygotuj_filtr()
    
    ' ========================================
    ' ========================================
    SIXP.LoadingFormModule.incLoadingForm
    
    ' 4
    wrksh.Cells(4, 1) = "TOTAL TOTAL"
    wrksh.Cells(4, 2) = inner_sum(wrksh.Range(wrksh.Cells(3, 2), wrksh.Cells(3, 100)))
    
    
    ' 5
    Set rng = wrksh.Cells(6, 1)
    rng.Offset(-1, 0) = "PPAP STATUS"
    Set rng = zrob_recursy_dla(fltr, m, rng, SIXP.ppap_status)
    
    SIXP.LoadingFormModule.incLoadingForm
    
    ' 10
    wrksh.Cells(10, 1) = "6P"
    wrksh.Cells(11, 1) = "DEL CONF, WHICH IS NOT CONNECTED WITH MRD PARAM."
    Set rng = wrksh.Cells(12, 1)
    Set rng = zrob_recursy_dla(fltr, m, rng, SIXP.Delivery_confirmation, E_SPEC_CASE_DO_NOT_TAKE_DEL_CONF_CONNECTED_WITH_MRD)
    
    SIXP.LoadingFormModule.incLoadingForm
    
    '15
    Set rng = wrksh.Cells(16, 1)
    rng.Offset(-1, 0) = "BEFORE OR ON/AFTER MRD"
    rng.Offset(-1, 2) = "MRD CW: "
    rng.Offset(-1, 3) = CStr(m_r_d)
    rng.Offset(-1, 4) = "BUILD START CW: "
    rng.Offset(-1, 5) = CStr(build_start)
    Set rng = zrob_recursy_dla(fltr, m, rng, SIXP.Delivery_confirmation, E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD_AND_AFTER_BUILD_START)
    
    
    SIXP.LoadingFormModule.incLoadingForm
    ' 20
    Set rng = wrksh.Cells(21, 1)
    rng.Offset(-1, 0) = "Del Conf"
    rng.Offset(-1, 1) = "MRD Date: "
    rng.Offset(-1, 2) = CStr(CDate(wez_date_mrd_z_details(d, sprawdz_czy_jest_sens_brac_date_mrd(d))))
    
    rng.Offset(-1, 3) = "MRD CW: "
    rng.Offset(-1, 4) = CStr(m_r_d)
    
    rng.Offset(-1, 5) = "BUILD START CW: "
    rng.Offset(-1, 6) = CStr(build_start)
    
    Set rng = zrob_recursy_dla(fltr, m, rng, SIXP.Delivery_confirmation)
    
    
    SIXP.LoadingFormModule.incLoadingForm
    
    ' 25
    Set rng = wrksh.Cells(26, 1)
    rng.Offset(-1, 0) = "Country Code"
    Set rng = zrob_recursy_dla(fltr, m, rng, SIXP.country_code)
    
    SIXP.LoadingFormModule.incLoadingForm
    
    ' 30
    Set rng = wrksh.Cells(31, 1)
    rng.Offset(-1, 0) = "CC Osea"
    Set rng = zrob_special_recursy_dla_cc_osea(fltr, m, rng, SIXP.country_code)
    
    SIXP.LoadingFormModule.incLoadingForm
    
    ' 35
    Set rng = wrksh.Cells(36, 1)
    rng.Offset(-1, 0) = "IN TRANSIT"
    rng.Offset(-1, 1) = "MRD: "
    rng.Offset(-1, 2) = CStr(CDate(wez_date_mrd_z_details(d, sprawdz_czy_jest_sens_brac_date_mrd(d))))
    rng.Offset(-1, 3) = "Today: "
    rng.Offset(-1, 4) = Date
    ' dzielimy po today, a nie po mrd
    Set rng = zrob_pus_recur(m, p, rng, True, Date)
    
    SIXP.LoadingFormModule.incLoadingForm
    
    
    ' ordered - po statusach
    ' 40
    Set rng = wrksh.Cells(41, 1)
    rng.Offset(-1, 0) = "Ordered"
    Set rng = zrob_recursy_dla(fltr, m, rng, SIXP.MRD1_Ordered_STATUS)
        
    
    ' Set puses = zrob_arkusz_puses(p, puses)
    
    
    'Columns("A:ZZ").Select
    'Selection.ColumnWidth = 12
    'Cells(1, 1).Select
    
    'wrksh.Select
    '
    'Columns("A:ZZ").Select
    'Selection.ColumnWidth = 12
    'Cells(1, 1).Select
    
    
    
    
    ' MsgBox "ready!" - za szybko
    '
    ''
    ' ======================================================
    
    SIXP.LoadingFormModule.incLoadingForm
    
    
    SIXP.LoadingFormModule.hideLoadingForm
End Sub


Private Function inner_sum(r As Range) As Long
    inner_sum = 0
    
    For Each ir In r
        If IsNumeric(ir) Then
            inner_sum = inner_sum + CLng(ir)
        End If
    Next ir
End Function

Private Function przygotuj_filtr() As String
    przygotuj_filtr = ";"
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")
    
    Do
        If CStr(r.Offset(0, 1)) = "1" Then
            przygotuj_filtr = przygotuj_filtr & CStr(r) & ";"
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function
    


Private Function zrob_arkusz_puses(mp As Worksheet, mpuses As Worksheet) As Worksheet

    'Public Enum E_PUS_SH
    '    O_INDX = 1
    '    O_PN
    '    O_DUNS
    '    O_FUP_code
    '    O_Pick_up_date
    '    O_Delivery_Date
    '    O_Pick_up_Qty
    '    O_PUS_Number
    'End Enum
    
    ' sekcja labelek
    With mpuses
        
        .Cells(1, 1) = "PN"
        .Cells(1, 2) = "DUNS"
        .Cells(1, 3) = "FUP CODE"
        .Cells(1, 4) = "PUS DATE"
        .Cells(1, 5) = "EDA"
        .Cells(1, 6) = "QTY"
        .Cells(1, 7) = "PUS #"
    End With
    
    Dim r As Range, f As Range
    Set f = mpuses.Cells(2, 1)
    
    Set r = przelicz_zasieg_dla_pusow(mp)
    
    Dim fst As Range
    
    Do
    
        Set fst = r.item(1)
        
        If CStr(Trim(fst)) <> "" Then
            
            For x = SIXP.O_PN To SIXP.O_PUS_Number
            
                If x = SIXP.O_Delivery_Date Or x = SIXP.O_Pick_up_date Then
                    f.Offset(0, x - SIXP.O_PN) = CDate(fst.Parent.Cells(fst.Row, x))
                Else
                    f.Offset(0, x - SIXP.O_PN) = CStr(fst.Parent.Cells(fst.Row, x))
                End If
                
            Next x
            
            Set f = f.Offset(1, 0)
        End If
        
        
        Dim tmp As Range
        Set tmp = r.item(2)
        If Trim(tmp) = "" Then
            Set tmp = fst.End(xlDown)
            
            If tmp.Row > r.item(r.Count).Row Then
                Set tmp = r.item(r.Count)
            End If
        End If
        Set r = r.Parent.Range(tmp, r.item(r.Count))
        
    Loop While r.Count > 1


    Set zrob_arkusz_puses = mpuses
End Function

Private Function zrob_pus_recur(m As Worksheet, mp As Worksheet, r As Range, czy_brac_bool_pod_date As Boolean, Optional d1 As Date) As Range
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Set dic = wypelnij_slownik_dla_pusow(m, dic, przelicz_zasieg_dla_pusow(mp))
    
    If czy_brac_bool_pod_date Then
        r = "RECV"
        r.Offset(0, 1) = "IN TRANSIT"
        r.Offset(0, 2) = "FUTURE"
        
        ' init by zeros
        r.Offset(1, 0) = 0
        r.Offset(1, 1) = 0
        r.Offset(1, 2) = 0
    End If
    
    
    For Each ki In dic.Keys
        If Trim(CStr(ki)) <> "" Then
            ' sekcja jak nie ma pustego
            ' ===================================
            If Not czy_brac_bool_pod_date Then
            
                r = ki
                r.Offset(1, 0) = dic.item(ki)
                
                Set r = r.Offset(0, 1)
                
            ElseIf czy_brac_bool_pod_date Then
            
                ' Debug.Print CDate(dic.item(ki).Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN))
                If CDate(dic.item(ki).range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) <= CDate(d1) Then
                    
                    r.Offset(1, 0) = r.Offset(1, 0) + 1
                Else
                
                
                    If CDate(dic.item(ki).range1.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) <= CDate(d1) Then
                        If CDate(dic.item(ki).range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) >= CDate(d1) Then
                            r.Offset(1, 1) = r.Offset(1, 1) + 1
                        End If
                    End If
                    
                    If CDate(dic.item(ki).range1.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) >= CDate(d1) Then
                        If CDate(dic.item(ki).range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) >= CDate(d1) Then
                            r.Offset(1, 2) = r.Offset(1, 2) + 1
                        End If
                    End If
                End If
            End If
            
            ' ===================================
        End If
    Next
End Function

Private Function przelicz_zasieg_dla_pusow(mp As Worksheet) As Range

    Set przelicz_zasieg_dla_pusow = _
        mp.Range(mp.Cells(2, SIXP.O_PN), mp.Cells(SIXP.POLOWA_CAPACITY_ARKUSZA, SIXP.O_PN))
End Function

Public Function zrob_special_recursy_dla_cc_osea(fltr As String, m As Worksheet, rng As Range, cc_column) As Range


    Dim eur_cc As Long
    Dim osea_cc As Long
    
    eur_cc = 0
    osea_cc = 0
    

    
    podlicz_osea eur_cc, osea_cc, fltr, przelicz_zasieg(m, SIXP.pn, cc_column)

    rng.Offset(0, 0) = "OSEA"
    rng.Offset(0, 1) = "EUR"

    rng.Offset(1, 0) = osea_cc
    rng.Offset(1, 1) = eur_cc

    Set zrob_special_recursy_dla_cc_osea = rng

End Function

Public Function zrob_recursy_dla(fltr As String, m As Worksheet, rng As Range, m_col, Optional e As E_SPECIAL_CASE_FOR_DEL_CONF) As Range
    
    Dim dic As Dictionary
    Set dic = New Dictionary
    
    Dim d As Worksheet
    Set d = m.Parent.Sheets("DETAILS")
    
    'If e = E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD Then
    '    Set dic = wypelnij_slownik(fltr, dic, _
    '        przelicz_zasieg(m, SIXP.pn, m_col), _
    '        E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD)
    'Else
    '    Set dic = wypelnij_slownik(fltr, dic, przelicz_zasieg(m, SIXP.pn, m_col))
    'End If
    
    If e = E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD Then
        Set dic = wypelnij_slownik(fltr, dic, _
            przelicz_zasieg(m, SIXP.pn, m_col), _
            E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD)
        
    ElseIf e = E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD_AND_AFTER_BUILD_START Then
        Set dic = wypelnij_slownik(fltr, dic, _
            przelicz_zasieg(m, SIXP.pn, m_col), _
            E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD_AND_AFTER_BUILD_START)
    Else
        Set dic = wypelnij_slownik(fltr, dic, przelicz_zasieg(m, SIXP.pn, m_col))
    End If
    
    For Each ki In dic.Keys
    

            
        If e = E_SPEC_CASE_DO_NOT_TAKE_DEL_CONF_CONNECTED_WITH_MRD Then
            
            If CStr(ki) <> "" And Not (CStr(ki) Like "*Y*CW*") Then
                rng = ki
                rng.Offset(1, 0) = iteruj_recur(fltr, 0, przelicz_zasieg(m, SIXP.pn, m_col), ki, E_EQUAL)
                
                Set rng = rng.Offset(0, 1)
            End If
        
        ElseIf e = E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD_AND_AFTER_BUILD_START Then
            
            
                ' tutaj sekcja ni bedzie miala *Y*CW poniewaz zostala z kluczy wykasowana w poprzedniej
                ' funkcji
                ' zalozenie jest takie ze wszystkie tutaj klucze biora udzial w zabawie nawet te puste poniewaz
                ' czyste mrd nie ma dodatkowego textu
                ' ======================================================
                ' ======================================================
                ''
                '
                rng = "BEFORE " & CStr(ki)
                rng.Offset(1, 0) = iteruj_recur(fltr, 0, _
                    przelicz_zasieg(m, SIXP.pn, m_col), _
                    przygotuj_my_pattern("BEFORE " & CStr(ki)), E_BEFORE_OR_AFTER_MRD_OR_AFTER_BUILD_START, d)
                
                Set rng = rng.Offset(0, 1)
                
                rng = "AFTER " & CStr(ki)
                rng.Offset(1, 0) = iteruj_recur(fltr, 0, _
                    przelicz_zasieg(m, SIXP.pn, m_col), _
                    przygotuj_my_pattern("AFTER " & CStr(ki)), _
                    E_BEFORE_OR_AFTER_MRD_OR_AFTER_BUILD_START, d)
                
                Set rng = rng.Offset(0, 1)
                
                
                rng = "AFTER BUILD START " & CStr(ki)
                rng.Offset(1, 0) = iteruj_recur(fltr, 0, _
                    przelicz_zasieg(m, SIXP.pn, m_col), _
                    przygotuj_my_pattern("AFTER BUILD START " & CStr(ki)), _
                    E_BEFORE_OR_AFTER_MRD_OR_AFTER_BUILD_START, d)
                
                Set rng = rng.Offset(0, 1)
                
                '
                ''
                ' ======================================================
                ' ======================================================
                
            
        Else
        
            If CStr(ki) <> "" Then
                rng = ki
                rng.Offset(1, 0) = iteruj_recur(fltr, 0, przelicz_zasieg(m, SIXP.pn, m_col), ki, E_EQUAL, d)
                
                Set rng = rng.Offset(0, 1)
            End If
            
        End If
    Next
    
    Set zrob_recursy_dla = rng
    
End Function

Private Function przygotuj_my_pattern(s As String) As String
    przygotuj_my_pattern = CStr(s)
End Function

Private Function wypelnij_slownik_dla_pusow(ByRef m As Worksheet, ByRef d As Dictionary, r As Range) As Dictionary
    
    Dim tr As TwoRanges
    
    Do
    
    
        If pn_on_master_and_is_still_under_resp(m, r.item(1)) Then
        
        
        
            Set tr = Nothing
            Set tr = New TwoRanges
        
            Set tr.range1 = r.item(1)
            Set tr.range2 = r.item(1)
        
            If CStr(tr.range1) <> "" Then
                If Not d.Exists(CStr(tr.range1)) Then
                    ' d.item(CStr(fst)) = CDate(fst.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN))
                    
                    ' to tutaj ponizej jest dziwne...
                    ' ----------------------------------
                    'Set d.item(CStr(tr.range1)) = tr
                    ' ----------------------------------
                    
                    
                    
                    d.Add CStr(tr.range1), tr
                Else
                    
                    'Debug.Print "d fro dic rng2: " & CDate(d.item(CStr(tr.range1)).range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN))
                    'Debug.Print "tr.rng2: " & CDate(tr.range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN))
                    
                    If CDate(d.item(CStr(tr.range1)).range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) < CDate(tr.range2.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) Then
                        Set d.item(CStr(tr.range1)).range2 = tr.range2
                    End If
                    
                    'Debug.Print CDate(d.item(CStr(tr.range1)).range1.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN))
                    'Debug.Print CDate(tr.range1.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN))
                    
                    If CDate(d.item(CStr(tr.range1)).range1.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) > CDate(tr.range1.Offset(0, SIXP.O_Delivery_Date - SIXP.O_PN)) Then
                        Set d.item(CStr(tr.range1)).range1 = tr.range1
                    End If
                End If
            End If
                
        End If
    
        Dim tmp As Range
        Set tmp = r.item(2)
        If Trim(tmp) = "" Then
            Set tmp = tmp.End(xlDown)
            
            If tmp.Row > r.item(r.Count).Row Then
                Set tmp = r.item(r.Count)
            End If
        End If
        
        ' Set d = wypelnij_slownik_dla_pusow(d, tail)
        
        
        
        Set r = r.Parent.Range(tmp, r.item(r.Count))
    Loop While r.Count > 1
    
    Set wypelnij_slownik_dla_pusow = d
    
End Function




Private Function pn_on_master_and_is_still_under_resp(ByRef m As Worksheet, ir As Range) As Boolean

    pn_on_master_and_is_still_under_resp = False
    
    
    Dim zasieg As Range
    Set zasieg = m.Range(m.Cells(2, SIXP.pn), m.Cells(SIXP.POLOWA_CAPACITY_ARKUSZA, SIXP.pn))
    
    Dim tmp As Range
    Set tmp = Nothing
    
    Set tmp = zasieg.Find(CStr(ir), LookIn:=xlValues, lookat:=xlWhole)
    
    If Not tmp Is Nothing Then
    
        If CStr(tmp) = CStr(ir) Then
        
            If sprawdz_resp_teraz(tmp) Then
                pn_on_master_and_is_still_under_resp = True
            Else
                pn_on_master_and_is_still_under_resp = False
            End If
        End If
    Else
        pn_on_master_and_is_still_under_resp = False
    End If
    
    
End Function

Private Function sprawdz_resp_teraz(ByRef r As Range) As Boolean
    sprawdz_resp_teraz = False
    
    Dim ir As Range
    Set ir = r.Parent.Cells(r.Row, SIXP.Responsibility)
    
    Dim resp As Range
    Set resp = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")
    Set resp = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range(resp, resp.End(xlDown))
    
    Dim tmp As Range
    Set tmp = resp.Find(CStr(ir), LookIn:=xlValues, lookat:=xlWhole)
    
    
    If Not tmp Is Nothing Then
        
        If CStr(tmp.Offset(0, 1)) = "1" Then
            sprawdz_resp_teraz = True
            Exit Function
        Else
            sprawdz_resp_teraz = False
        End If
    Else
        sprawdz_resp_teraz = False
    End If
    
    
    
End Function


Private Sub podlicz_osea(ByRef eur As Long, ByRef osea As Long, fltr As String, r As Range)


    ' r ===> from cc column
    
    Dim fst As Range
    Do
        Set fst = r.item(1)
        ' If fst.Parent.Cells(fst.Row, SIXP.Responsibility) Like "*" & fltr & "*" Then
        If sprawdz_resp_teraz(fst) Then
        
        
            If sprawdz_czy_osea(fst) Then
                osea = osea + 1
            Else
                eur = eur + 1
            End If
        End If
        
        If r.Count > 1 Then
            Set r = r.Parent.Range(r.item(2), r.item(r.Count))
        Else
            Exit Do
        End If
        
    Loop While True
    
End Sub

Private Function sprawdz_czy_osea(s) As Boolean
    sprawdz_czy_osea = False
    
    Dim ccsh As Worksheet
    Set ccsh = ThisWorkbook.Sheets(SIXP.G_CC_SH_NM)
    
    Dim ccr As Range
    Set ccr = ccsh.Range("B1")
    Do
    
        If UCase(CStr(Trim(s))) = UCase(CStr(Trim(ccr))) Then
            If CLng(ccr.Offset(0, 3)) = CLng(1) Then
                sprawdz_czy_osea = True
                Exit Function
            Else
                Exit Function
            End If
        End If
    
        Set ccr = ccr.Offset(1, 0)
    Loop Until Trim(ccr) = ""
    
End Function


Private Function wypelnij_slownik(fltr As String, ByRef d As Dictionary, r As Range, Optional e As E_SPECIAL_CASE_FOR_DEL_CONF) As Dictionary
    
    Dim fst As Range
    
    Do
    
        Set fst = r.item(1)
        'Debug.Print fst.Parent.Cells(fst.Row, SIXP.Responsibility)
        ' to jest juz troche nieaktualne poniewaz nie bierze pod uwage danych ktore sami wczesniej zdefiniowalismy
        ' If fst.Parent.Cells(fst.Row, SIXP.Responsibility) Like "*" & fltr & "*" Then
        ' If (fltr = "*") Or (sprawdz_resp_teraz(fst.Parent.Cells(fst.Row, SIXP.Responsibility)) And fltr = "") Then
        ' If CStr(fltr) Like "*" & CStr(fst.Parent.Cells(fst.Row, SIXP.Responsibility)) & "*" Then
        ' Debug.Print CStr(fst.Parent.Cells(fst.Row, SIXP.Responsibility))
        
        If (Not Application.WorksheetFunction.IsNA(fst)) Or (Not Application.WorksheetFunction.IsError(fst)) Then
        
            If CStr(fltr) = "*" Or CStr(fltr) Like "*;" & CStr(fst.Parent.Cells(fst.Row, SIXP.Responsibility)) & ";*" Then
                
                If e = E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD Then
                
                    ' tutaj warunek okrjony sprawdzajacy tylko czy dany element zawiera Y*CW
                    If CStr(fst) Like "*Y*CW*" Then
                        
                        
                        ' sekcja wyodrebniajaca before and after dla wybranych del confow.
                        
                        ' jednak to nie uwzglednia czystego {MRD}
                        ' to jest czyste MRD
                        If Left(CStr(fst), 1) = "Y" Then
                            If Not d.Exists("MRD") Then
                                d.Add "MRD", Nothing
                            End If
                        Else
                            If Not d.Exists(CStr(Split(CStr(fst), " ")(0))) Then
                                d.Add CStr(Split(CStr(fst), " ")(0)), Nothing
                            End If
                        End If
                    End If
                ElseIf e = E_SPEC_CASE_COUNT_BEFORE_AND_AFTER_MRD_AND_AFTER_BUILD_START Then
                    
                    If CStr(fst) Like "*Y*CW*" Then
                    
                        ' sekcja wyodrebnia 3 mozliwosci:
                        ' 1. przed mrd
                        ' 2. po mrd ale przed build start
                        ' 3. po build start
                        
                        If Left(CStr(fst), 1) = "Y" Then
                            If Not d.Exists("MRD") Then
                                d.Add "MRD", Nothing
                            End If
                        Else
                            If Not d.Exists(CStr(Split(CStr(fst), " ")(0))) Then
                                d.Add CStr(Split(CStr(fst), " ")(0)), Nothing
                            End If
                        End If
                        
                    End If
                Else
        
                    If Not d.Exists(CStr(fst)) Then
                        d.Add CStr(fst), Nothing
                    End If
                End If
                        
            End If
        End If
                    
        Set r = r.Parent.Range(r.item(2), r.item(r.Count))
        
    Loop While r.Count > 1
    
    Set wypelnij_slownik = d
    
End Function

Public Function przelicz_zasieg(m As Worksheet, col1, docelowa_kolumna) As Range

    If Trim(m.Cells(2, col1)) <> "" Then
    
    
        If Int(docelowa_kolumna) = Int(Delivery_confirmation) Then
            ostatni_wiersz = m.Cells(SIXP.POLOWA_CAPACITY_ARKUSZA, docelowa_kolumna).End(xlUp).Row + 10
        Else
            ostatni_wiersz = m.Cells(SIXP.POLOWA_CAPACITY_ARKUSZA, docelowa_kolumna).End(xlUp).Row
        End If
    
        Set przelicz_zasieg = _
            m.Range(m.Cells(2, docelowa_kolumna), m.Cells(ostatni_wiersz, docelowa_kolumna))
    Else
        Set przelicz_zasieg = m.Cells(2, docelowa_kolumna)
    End If
    

End Function

Public Function iteruj_recur(fltr As String, start, r As Range, filter, e As E_MATCH, Optional d As Worksheet) As Long
    
    ' robimy rekurencje - pobierz pierwszy element zasiegu
    ' i zostaw reszte dla kolejnej rekurencji
    
    ' Optional d As Worksheet for DETAILS
    
    
    ' ????
    ' start = 0
    
    Dim fst As Range
    
    
    Do
    
        Set fst = r.item(1)
        
        If Trim(fst.Value) <> "" Then
            
            ' If CStr(fltr) Like "*" & CStr(fst.Parent.Cells(fst.Row, SIXP.Responsibility)) & "*" Then
            'If (fltr = "*") Or (sprawdz_resp_teraz(fst.Parent.Cells(fst.Row, SIXP.Responsibility)) And fltr = "") Then
            ' po krotce pierwszy filtr jest pod respa
            ' drugi filtr jest pod dedykowana kolumne
            ' Debug.Print CStr(fst.Parent.Cells(fst.Row, SIXP.Responsibility))
            If CStr(fltr) = "*" Or CStr(fltr) Like "*;" & CStr(fst.Parent.Cells(fst.Row, SIXP.Responsibility)) & ";*" Then
                
                If e = E_LIKE Then
                    If fst Like "*" & CStr(filter) & "*" Then
                        start = start + 1
                    End If
                ElseIf e = E_EQUAL Then
                    If CStr(fst) = CStr(filter) Then
                        start = start + 1
                    End If
                ElseIf e = E_NOT_EQUAL Then
                    If CStr(fst) <> CStr(filter) Then
                        start = start + 1
                    End If
                ElseIf e = E_BEFORE_OR_AFTER_MRD Then
                
                
                    ' najwygodniej zaczac od tego co wiem napewno
                    ' wez z arkusza details wartosc mrd poniewaz na jej bazie bede decydowal jakie del confy sa cacy a jakie nie
                    date_mrd = wez_date_mrd_z_details(d, _
                        sprawdz_czy_jest_sens_brac_date_mrd(d))
                        
                    
                    If porownaj_daty_zafiltruj_i_okresl_czy_dajemy_plus_one(CDate(date_mrd), _
                        wez_date_z_del_conf_param(CStr(fst)), _
                        filter, _
                        fst) _
                            Then
                                start = start + 1
                    End If
                    
                ElseIf e = E_BEFORE_OR_AFTER_MRD_OR_AFTER_BUILD_START Then
                
                    
                    date_mrd = wez_date_mrd_z_details(d, _
                        sprawdz_czy_jest_sens_brac_date_mrd(d))
                        
                    date_build_start = wez_date_build_start_z_details(d)
                    date_bom = wez_date_bom_freeze_z_details(d)
                    
                    date_del_conf = wez_date_z_del_conf_param(CStr(fst))
                    
                    If porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one( _
                        CDate(date_mrd), CDate(date_build_start), _
                        CDate(date_del_conf), filter, fst) _
                        Then
                            start = start + 1
                    End If

                End If

            End If
        
        End If
        
        If r.Count > 1 Then
            Set r = r.Parent.Range(r.item(2), r.item(r.Count))
        Else
            Exit Do
        End If
        
    Loop While True
    
    iteruj_recur = start
    
    
End Function

Private Function porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one( _
    mrd_date As Date, build_start_date As Date, _
    del_conf_monday_date As Date, _
    str_filter, _
    r As Range) _
        As Boolean
        
        
            porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one = False
            
            
            
        
            Dim delConfFromR As String
        
            If CDate(del_conf_monday_date) <> CDate("1900-01-01") Then

                delConfFromR = CStr(Split(r, " ")(0))
                
                If delConfFromR Like "Y*CW*" Then
                    delConfFromR = "MRD"
                End If
                
                
                stripped_str_filter = Trim(Replace( _
                    Replace( _
                        Replace( _
                            CStr(str_filter), "BEFORE", "") _
                        , "AFTER", "") _
                    , "BUILD START", ""))

                If Trim(CStr(stripped_str_filter)) = Trim(CStr(delConfFromR)) Then
            
                    If CStr(str_filter) Like "*BEFORE*" Then
                        
                        If CDate(mrd_date) >= CDate(del_conf_monday_date) Then
                            porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one = True
                        End If
                    ElseIf CStr(str_filter) Like "*AFTER*" Then
                    
    
                        If CStr(str_filter) Like "*AFTER*" And (Not CStr(str_filter) Like "*AFTER BUILD START*") Then
                    
                            If CDate(mrd_date) < CDate(del_conf_monday_date) And CDate(build_start_date) >= CDate(del_conf_monday_date) Then
                                porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one = True
                            End If
                            
                            
                        ElseIf CStr(str_filter) Like "*AFTER BUILD START*" Then
                            
                            If CDate(build_start_date) < CDate(del_conf_monday_date) Then
                                porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one = True
                            End If
                        End If
                    End If
                    
                End If
                
            Else
                porownaj_daty_mrd_i_build_start_potem_zafiltruj_i_okresl_czy_dajemy_plus_one = False
            End If
        
End Function


Private Function porownaj_daty_zafiltruj_i_okresl_czy_dajemy_plus_one(mrd_date As Date, _
    del_conf_monday_date As Date, _
    str_filter, _
    r As Range) _
        As Boolean
        
            ' czyli jako tako data zostala odnaleziona
            If CDate(del_conf_monday_date) <> CDate("1900-01-01") Then
                
                If (CStr(r) Like "*" & CStr(str_filter) & "*") Or (CStr(str_filter) Like "*MRD*") Then
        
                    If CStr(str_filter) Like "*BEFORE*" Then
                        
                        If CDate(mrd_date) >= CDate(del_conf_monday_date) Then
                            porownaj_daty_zafiltruj_i_okresl_czy_dajemy_plus_one = True
                        End If
                    ElseIf CStr(str_filter) Like "*AFTER*" Then
                    
                        If CDate(mrd_date) < CDate(del_conf_monday_date) Then
                            porownaj_daty_zafiltruj_i_okresl_czy_dajemy_plus_one = True
                        End If
                    End If
                
                End If
            Else
                porownaj_daty_zafiltruj_i_okresl_czy_dajemy_plus_one = False
            End If
    
    
End Function

Private Function wez_date_z_del_conf_param(s As String) As Date
    
    ' na poczatku s jest pelnym textem z del conf  - nalezy sciagnac zbedne dane
    If CStr(s) Like "*Y*CW*" Then
        ' lecimy dalej
        
        ' take only ycw part
        ycw = zrob_y_cw(s)
        
        wez_date_z_del_conf_param = parsuj_y_cw_do_daty_poniedzialkowej_arg_as_str(CStr(ycw))
        
    Else
        ' tym sie w ogole nie zajmujemy
        wez_date_z_del_conf_param = CDate("1900-01-01")
    End If
End Function


Private Function zrob_y_cw(s As String) As String
    
    tmp = s
    If Left(s, 1) = "Y" Then
    Else
        tmp = zrob_y_cw(Mid(s, 2, Len(s) - 1))
    End If
    
    zrob_y_cw = tmp
    
End Function



Private Function wez_date_mrd_z_details(details_sh As Worksheet, directly_date_or_parse_from_str_mrd As Boolean) As Date
    
    If directly_date_or_parse_from_str_mrd Then
        wez_date_mrd_z_details = CDate(Format(details_sh.Cells(SIXP.E_MRD_DATE, 2), "yyyy-mm-dd"))
    Else
        wez_date_mrd_z_details = CDate(parsuj_y_cw_do_daty_poniedzialkowej(details_sh.Cells(SIXP.mrd, 2)))
    End If
    
    
End Function

Private Function wez_date_build_start_z_details(details_sh As Worksheet) As Date

    wez_date_build_start_z_details = CDate(parsuj_y_cw_do_daty_poniedzialkowej(details_sh.Cells(SIXP.build_start, 2)))
End Function


Private Function wez_date_bom_freeze_z_details(details_sh As Worksheet) As Date
        
    wez_date_bom_freeze_z_details = CDate(parsuj_y_cw_do_daty_poniedzialkowej(details_sh.Cells(SIXP.BOM, 2)))
End Function

Private Function parsuj_y_cw_do_daty_poniedzialkowej(r As Range) As Date
    ' sekcja parsu - r to komorka zawierajaca text y cw
    
    If CStr(r) Like "Y*CW*" Then
        
        y = Mid(CStr(r), 2, 4)
        d_str = y & "-01-01"
        Dim d As Date
        d = CDate(d_str)
        
        Do
            cw = CLng(Right(CStr(r), Len(CStr(r)) - 7))
            
            If CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(d))) = CLng(cw) Then
                parsuj_y_cw_do_daty_poniedzialkowej = d
                Exit Do
            End If
            d = d + 1
        Loop While CLng(Year(CDate(d_str))) = CLng(y)
    Else
        MsgBox "parametr MRD jest zle zdefiniowany"
        End
    End If
End Function

Private Function parsuj_y_cw_do_daty_poniedzialkowej_arg_as_str(r As String) As Date
    ' sekcja parsu - r to komorka zawierajaca text y cw
    
    If CStr(r) Like "Y*CW*" Then
        
        y = Mid(CStr(r), 2, 4)
        d_str = y & "-01-01"
        Dim d As Date
        d = CDate(d_str)
        
        Do
            cw = CLng(Right(CStr(r), Len(CStr(r)) - 7))
            
            If CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(d))) = CLng(cw) Then
                parsuj_y_cw_do_daty_poniedzialkowej_arg_as_str = CDate(d)
                Exit Do
            End If
            d = d + 1
        Loop While CLng(Year(CDate(d_str))) = CLng(y)
    Else
        MsgBox "parametr MRD jest zle zdefiniowany"
        End
    End If
End Function

Private Function sprawdz_czy_jest_sens_brac_date_mrd(details_sh As Worksheet) As Boolean
    If IsDate(details_sh.Cells(SIXP.E_MRD_DATE, 2)) Then
        sprawdz_czy_jest_sens_brac_date_mrd = True
    Else
        sprawdz_czy_jest_sens_brac_date_mrd = False
    End If
    
End Function

Private Sub fill_wiz_buff_with_all_details_but_transpose_from_o1(details_sh As Worksheet, wiz_buff_sh As Worksheet)

    Dim rw As Range, rd As Range
    Set rw = wiz_buff_sh.Range("O1")
    
    
    ' transposed data from details
    For x = SIXP.plt To E_UNIQUE_ID
        rw.Offset(0, x - 1) = details_sh.Cells(x, 2)
    Next x
    

End Sub
