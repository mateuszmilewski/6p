VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewProj 
   Caption         =   "Projekt"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6345
   OleObjectBlob   =   "NewProj.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClear_Click()

    Me.TextBoxCW = ""
    Me.TextBoxFaza = ""
    Me.TextBoxPlt = ""
    Me.TextBoxProj = ""
    
    Me.DTPicker1 = Date
    Me.TextBoxCW = SIXP.GlobalFooModule.parse_from_date_to_yyyycw(Date)
End Sub

Private Sub BtnDelete_Click()



    ans = MsgBox("Czy jestes pewien tego, co robisz?", vbOKCancel, "Delete item prompt")


    If ans = vbYes Then
        Dim m As Worksheet, r As Range
        Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
        
        If ThisWorkbook.ActiveSheet.Name = m.Name Then
        
            ' ==================================================
            Set r = validate_and_then_go_to_active_cell
            ' ==================================================
            
            If r Is Nothing Then
                MsgBox "Akcja nie jest dozwolona!"
            Else
                ' to jest akcja z edycji
                ' =================================================
                ' =================================================
                'r.Value = Me.TextBoxProj
                'r.Offset(0, 1).Value = Me.TextBoxPlt
                'r.Offset(0, 2).Value = Me.TextBoxFaza
                'r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
                'r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
                ' =================================================
                ' =================================================
                
                Dim dm As DeletionManager
                Set dm = New DeletionManager
                
                
                dm.usun_kazde_wystapienie_dla_aktywnej_komorki r
                
                Set dm = Nothing
                
                ' =================================================
                ' =================================================
            End If
        Else
            ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
            MsgBox "nie mozna wykonac akcji w tej lokalizacji pliku - makro samo Cie przesunelo na wlasciwy arkusz."
        End If
    Else
        MsgBox "nic sie nie stalo!"
    End If
        
End Sub

Private Sub BtnEdit_Click()
    
    Dim m As Worksheet, r As Range
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    
    If ThisWorkbook.ActiveSheet.Name = m.Name Then
    
    
        ' ==================================================
        Set r = validate_and_then_go_to_active_cell
        ' ==================================================
        
        If r Is Nothing Then
            MsgBox "Akcja nie jest dozwolona!"
        Else
        
            If r.Row > 1 Then
        
                r.Value = Me.TextBoxProj
                r.Offset(0, 1).Value = Me.TextBoxPlt
                r.Offset(0, 2).Value = Me.TextBoxFaza
                r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
                r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
            Else
                MsgBox "chcesz podmienic nazwy kolumn! Nie jest to dozwolona akcja!"
            End If
        End If
    Else
        ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
        MsgBox "nie mozna wykonac akcji w tej lokalizacji pliku - makro samo Cie przesunelo na wlasciwy arkusz."
    End If
End Sub

Private Sub BtnGoToDetails_Click()
    Hide
    Dim l As T_Link
    Set l = New T_Link
    Dim lr As Linker
    Set lr = New Linker
    l.zrob_mnie_z_argsow Me.TextBoxProj, Me.TextBoxPlt, Me.TextBoxFaza, Me.TextBoxCW
    run_FormMain CStr(lr.return_full_concated_r_string_comma_seperated(l))
End Sub

Private Sub BtnImport_Click()
    ' funkcja importu - calkiem wazna
    ' ---------------------------------------------------------
    
    ' wczesniej ten msgbox mial byc jako tako masowy
    ' jednak z perspektywy designu calej apki nie moge tak zrobic
    ' zatem zatem: tutaj tylko dodaje dane do tego formularza
    'ans = MsgBox("Czy chcesz zaciagnac jednorazowo informacje z pliku Wizard?", vbOKCancel, "Wizard Synchro")
    
    'If ans = vbOK Then
        Hide
        FormCatchWizard.ListBox1.Clear
        
        For Each w In Workbooks
            With FormCatchWizard.ListBox1
                .AddItem w.Name
            End With
        Next w
        
        FormCatchWizard.Show
    'End If
    
    ' ---------------------------------------------------------
End Sub

Private Sub BtnSubmit_Click()
    ' tutaj dodajemy nowy projekt na spod w arkuszu main
    
    Dim m As Worksheet
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    Dim r As Range
    ' oprocz tego co pisze w nazwie funkcji dodatkowo sprawdza te same projekty
    ' z roznymi cw i podputyuje co z tym fantem zrobic
    ' jesli chodzi o status to nie ma znaczenia
    Set r = validate_and_then_go_to_first_empty_cell(m)
    
    r.Value = Me.TextBoxProj
    r.Offset(0, 1).Value = Me.TextBoxPlt
    r.Offset(0, 2).Value = Me.TextBoxFaza
    r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
    r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
End Sub



Private Function validate_and_then_go_to_first_empty_cell(ByRef m As Worksheet) As Range


    ' ten sub oprocz swojej nazwy i podpisanej funkcjonalnosci posiada jeszcze mozliwosc podjecia decyzji
    ' o zaktualizowaniu juz istniejacego projektu
    ' o inny CW
    
    Dim r As Range
    Set r = m.Cells(1, 1)
    Do
        If CStr(Me.TextBoxProj) = Trim(CStr(r)) Then
            If CStr(Me.TextBoxPlt) = Trim(CStr(r.Offset(0, 1))) Then
                If Trim(CStr(r.Offset(0, 2).Value)) = Trim(CStr(Me.TextBoxFaza)) Then
                    If Trim(CStr(r.Offset(0, 3).Value)) = CStr(Me.TextBoxCW) Then
                        ans = MsgBox("duplikat! masz inny status? chcesz go podmienic?", vbYesNo)
                        If ans = vbYes Then
                            Set validate_and_then_go_to_first_empty_cell = r
                            Exit Function
                            
                        Else
                            MsgBox "apka konczy dzialanie"
                            End
                        End If
                    Else
                        ans = MsgBox("Projekt z nowym CW, czy chcesz go podmienic!", vbYesNo)
                        
                        If ans = vbYes Then
                        
                            Set validate_and_then_go_to_first_empty_cell = r
                            Exit Function
                        
                        ElseIf ans = vbNo Then
                            
                            ans = MsgBox("Czy chcesz zatem dodac ten sam projekt na dole z nowa data?", vbYesNo)
                            
                            If ans = vbYes Then
                                Set r = r.End(xlDown).End(xlDown).End(xlUp).Offset(1, 0)
                                Exit Do
                            Else
                                MsgBox "Logika konczy dzialanie bez wykonanej akcji na danych"
                                End
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Set validate_and_then_go_to_first_empty_cell = r
End Function

Private Function validate_and_then_go_to_active_cell() As Range
    
    Dim r As Range
    Set r = ActiveCell
    Set r = r.Parent.Cells(r.Row, 1)
    Do
        If Trim(CStr(r)) <> "" _
            Or Trim(CStr(r.Offset(0, 1))) <> "" _
            Or Trim(CStr(r.Offset(0, 2).Value)) <> "" _
            Or Trim(CStr(r.Offset(0, 3).Value)) <> "" Then
            
            Set validate_and_then_go_to_active_cell = r
            Exit Function
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Set validate_and_then_go_to_active_cell = Nothing
End Function

Private Sub BtnZduplikuj_Click()


    ans = MsgBox("Czy napewno chcesz stworzyc zduplikowana informacje opatrzona tylko nowym CW?", vbOKCancel, "NEW ITEM")
    
    If ans = vbOK Then
    
        If CStr(Me.TextBoxSelectedCW) <> CStr(Me.TextBoxCW) Then
    
            Dim m As Worksheet
            Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
            
            Dim r As Range
            ' oprocz tego co pisze w nazwie funkcji dodatkowo sprawdza te same projekty
            ' z roznymi cw i podputyuje co z tym fantem zrobic
            ' jesli chodzi o status to nie ma znaczenia
            Set r = validate_and_then_go_to_first_empty_cell(m)
            
            r.Value = Me.TextBoxProj
            r.Offset(0, 1).Value = Me.TextBoxPlt
            r.Offset(0, 2).Value = Me.TextBoxFaza
            r.Offset(0, 3).Value = CLng(Me.TextBoxCW)
            r.Offset(0, 4).Value = Me.ComboBoxStatus.Value
            
            ' tl_old i new roznica sie data
            Dim tl_old As T_Link, tl_new As T_Link
            Set tl_old = New T_Link
            Set tl_new = New T_Link
            tl_old.zrob_mnie_z_argsow Me.TextBoxProj, Me.TextBoxPlt, Me.TextBoxFaza, Me.TextBoxSelectedCW
            tl_new.zrob_mnie_z_argsow Me.TextBoxProj, Me.TextBoxPlt, Me.TextBoxFaza, Me.TextBoxCW
            dane_dla_nowego_skopiuj_ze_starego tl_old, tl_new, m
            
        Else
        
            MsgBox "probujesz zrobic duplikat z ta sama data - uzyj przycisku edytuj!"
        End If
    End If
End Sub

Private Sub dane_dla_nowego_skopiuj_ze_starego(old_tl As T_Link, new_tl As T_Link, m As Worksheet)

    ' lecimy przez wszystkie arkusza zawierajaca dane
    ' nie wrzucaj z gory!
    'For x = SIXP.e_main_last_update_on_order_release_status - 1 To SIXP.e_main_last_update_on_resp - 1
    '    r.Offset(0, x) = new_tl.cw
    'Next x
    
    work_on_ SIXP.G_order_release_status_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_order_release_status - 1)
    
    work_on_ SIXP.G_recent_build_plan_changes_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_recent_build_plan_changes - 1)
    work_on_ SIXP.G_cont_pnoc_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_chart_contracted_pnoc - 1)
    work_on_ SIXP.G_osea_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_osea - 1)
    
    work_on_ SIXP.G_totals_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_totals - 1)
    work_on_ SIXP.G_xq_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_xq - 1)
    work_on_ SIXP.G_del_conf_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_del_conf - 1)
    work_on_ SIXP.G_open_issues_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_open_issues - 1)
    
    work_on_ SIXP.G_resp_sh_nm, old_tl, new_tl, m, CLng(SIXP.e_main_last_update_on_resp - 1)
        
        
    
End Sub

Private Sub work_on_(str_sh_nm As String, old_tl As T_Link, new_tl As T_Link, m As Worksheet, offset_w_sh_main As Long)
    
    Dim sh As Worksheet
    Dim l As Linker
    Set l = New Linker
    Dim old_r As Range, new_r As Range
    
    Dim ostatnia_kolumna_arkusza As Long
    
    Set sh = ThisWorkbook.Sheets(str_sh_nm)
    Set new_r = pierwszy_pusty_wiersz(sh)
    ostatnia_kolumna_arkusza = okresl_ostatnia_kolumne(sh)
    Set old_r = old_tl.znajdz_siebie_w_arkuszu(sh)
    
    If Not old_r Is Nothing Then
    
    
        ' main!
        ' ============================================
        Dim r As Range
        Set r = new_tl.znajdz_siebie_w_arkuszu(m)
        r.Offset(0, offset_w_sh_main) = new_tl.cw
        ' ============================================
        
        
        ' skopiuj_stare_dane_do_nowego old_r, new_r, old_tl, new_tl
        ' zrobimy to tutaj nie bedziemy znowu wydzielac procedury - bez sensu
        ' przy tej ilosci argsow
        ' =====================================================================
        
        
        ' zostawiamy czesc zwiazana z linkowaniem - przepiszemy sobie dane z obiektow
        ' typu T_Link
        new_r.Offset(0, 0) = new_tl.project
        new_r.Offset(0, 1) = new_tl.plt
        new_r.Offset(0, 2) = new_tl.faza
        new_r.Offset(0, 3) = new_tl.cw
        
        ' e_link_cw == 4 zatem pasuje z racji offsetowania!
        ' -----------------------------------------------------
        For x = SIXP.e_link_cw To ostatnia_kolumna_arkusza - 1
            new_r.Offset(0, x) = old_r.Offset(0, x)
        Next x
        ' -----------------------------------------------------
    
        ' =====================================================================
    Else
        
    End If
        
End Sub


Private Function pierwszy_pusty_wiersz(s As Worksheet) As Range
    
    Dim r As Range
    Set r = s.Cells(1, 1)
    Do
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Set pierwszy_pusty_wiersz = r
End Function

Private Function okresl_ostatnia_kolumne(s As Worksheet) As Long
    okresl_ostatnia_kolumne = 1
    
    Dim r As Range
    Set r = s.Cells(1, 1)
    Do
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
    
    okresl_ostatnia_kolumne = CLng(r.Column)
End Function

Private Sub ComboBoxFAZA_Change()
    Me.TextBoxFaza = CStr(Me.ComboBoxFAZA.Value)
End Sub

Private Sub ComboBoxPLT_Change()
    Me.TextBoxPlt = CStr(Me.ComboBoxPLT.Value)
End Sub



Private Sub DTPicker1_Change()
    Me.TextBoxCW = SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPicker1.Value))
End Sub
