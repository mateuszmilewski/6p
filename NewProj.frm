VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewProj 
   Caption         =   "Projekt"
   ClientHeight    =   8040
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

Private Sub BtnClear_Click()

    Me.TextBoxCW = ""
    Me.TextBoxFaza = ""
    Me.TextBoxPlt = ""
    Me.TextBoxProj = ""
    
    Me.DTPicker1 = Date
    Me.TextBoxCW = SIXP.GlobalFooModule.parse_from_date_to_yyyycw(Date)
End Sub

Private Sub BtnCopy_Click()

    If Trim(Me.TextBoxProj.Value) <> "" Then
    
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
    
    
            skopiuj_dane_z_innego_projektu CStr(Me.TextBoxProj), CStr(Me.TextBoxPlt), CStr(Me.TextBoxFaza), CStr(Me.TextBoxCW), CStr(Me.ComboBoxStatus)
        End If
    End If
End Sub


Private Sub skopiuj_dane_z_innego_projektu(proj, plt, faza, cw, status)
    
    Hide
    
    Dim myNewLink As T_Link
    Set myNewLink = New T_Link
    
    myNewLink.zrob_mnie_z_argsow proj, plt, faza, cw
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    Dim r As Range
    Set r = sh.Range("A2")
    
    GetProject.ListBoxProjects.Clear
    GetProject.ListBoxPLT.Clear
    GetProject.ListBoxFaza.Clear
    GetProject.ListBoxCW.Clear
    
    
    Do
        ' --------------------------------------------------------------------------------------------------------------------
        Set SIXP.GetProject.newLink = myNewLink
        ' SIXP.GetProject.ListBoxProjects.AddItem qinnercncat(r, r.Offset(0, 1), r.Offset(0, 2), r.Offset(0, 3), r.Offset(0, 4))
        GetProject.ListBoxProjects.AddItem r
        GetProject.ListBoxPLT.AddItem r.Offset(0, 1)
        GetProject.ListBoxFaza.AddItem r.Offset(0, 2)
        GetProject.ListBoxCW.AddItem r.Offset(0, 3)
        ' --------------------------------------------------------------------------------------------------------------------
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    SIXP.GetProject.Show
    
End Sub

Private Function qinnercncat(r, r01, r02, r03, r04)
    
    qinnercncat = "" & r & ";" & r01 & ";" & r02 & ";" & r03 & ";" & r04
End Function

Private Sub BtnDelete_Click()



    ans = MsgBox("Czy jestes pewien tego, co robisz?", vbOKCancel, "Delete item prompt")


    If ans = vbYes Then
        Dim m As Worksheet, r As Range
        Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
        
        If ThisWorkbook.ActiveSheet.name = m.name Then
        
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
    
    
    If ThisWorkbook.ActiveSheet.name = m.name Then
    
    
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


    If Trim(Me.TextBoxProj.Value) = "" Then
        MsgBox "Brak nazwy projektu!"
    Else

        Hide
        Dim l As T_Link
        Set l = New T_Link
        Dim lr As Linker
        Set lr = New Linker
        l.zrob_mnie_z_argsow Me.TextBoxProj, Me.TextBoxPlt, Me.TextBoxFaza, Me.TextBoxCW
        run_FormMain CStr(lr.return_full_concated_r_string_comma_seperated(l))
    
    End If
End Sub

Private Sub BtnImport_Click()
    ' funkcja importu - calkiem wazna
    ' ---------------------------------------------------------
    
    ' wczesniej ten msgbox mial byc jako tako masowy
    ' jednak z perspektywy designu calej apki nie moge tak zrobic
    ' zatem zatem: tutaj tylko dodaje dane do tego formularza
    ans = MsgBox("Czy chcesz zaciagnac jednorazowo informacje z otwartego pliku Wizard?", vbOKCancel, "Wizard Synchro")
    
    If ans = vbOK Then
        Hide
        
        
        ' usuniecie danych z wizard buff
        ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("a1:zz1000").Clear
        
        FormCatchWizard.ListBox1.Clear
        FormCatchWizard.ListBox1.MultiSelect = fmMultiSelectSingle
        
        For Each w In Workbooks
            With FormCatchWizard.ListBox1
                .AddItem w.name
            End With
        Next w
        FormCatchWizard.czy_start_pochodzi_z_open_issues = False
        FormCatchWizard.BtnImportOpenIssues.Enabled = False
        FormCatchWizard.BtnJustImport.Enabled = True
        FormCatchWizard.BtnSubmit.Enabled = True
        FormCatchWizard.Show
    Else
        MsgBox "logika zatrzymana"
    End If
    
    ' ---------------------------------------------------------
End Sub

Private Sub BtnImportWizBuff_Click()
        
    Dim wh As WizardHandler
    Set wh = New WizardHandler

        
    With Me
        .TextBoxCW = wh.get_cw()
        .TextBoxFaza = wh.get_faza_from_buffer()
        .TextBoxPlt = wh.get_plt_from_buffer()
        .TextBoxProj = wh.get_proj_from_buffer() & " " & wh.get_biw_ga_from_buffer & " " & " MY: " & wh.get_my_from_buffer()
        .ComboBoxStatus = SIXP.GlobalCrossTriangleCircleModule.putCross
        
        
        
        ' zmiana od wersji 0.26
        ' 2017-01-24
        ' -----------------------------------------------
        '
        ' odblokowuje opcje zaciagania z buffa ale musi jeszcze user dokliknac
        .CheckBoxWizardContent.Enabled = True
        .CheckBoxWizardContent.Value = False
        
    End With
        
    Set wh = Nothing
End Sub

Private Sub BtnSubmit_Click()
    ' tutaj dodajemy nowy projekt na spod w arkuszu main
    
    If Me.TextBoxProj <> "" And Me.TextBoxPlt <> "" And Me.TextBoxFaza.Value <> "" Then
    
    
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
        
        If Me.CheckBoxWizardContent.Value Then
        
            ' zbieramy dodatkowo info z buffa
            ' ---------------------------------------------------------------------------
            ' MsgBox "not implemented yet!"
            If sprawdzCzyMaSensZaciagacDaneZWizardBuff() Then
                doMassImport r
            Else
                MsgBox "nie ma czego importowac z wizard buffer!"
            End If
            ' ---------------------------------------------------------------------------
        End If
    
    
    Else
        ' no project at all pls fill data
        MsgBox "no input data!"
    End If
End Sub


Private Function sprawdzCzyMaSensZaciagacDaneZWizardBuff() As Boolean


    sprawdzCzyMaSensZaciagacDaneZWizardBuff = False
    sprawdzCzyMaSensZaciagacDaneZWizardBuff = checkWizardBufferCells()
    
End Function

Private Function checkWizardBufferCells() As Boolean
    
    checkWizardBufferCells = False
    tmp = False
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    Dim r As Range
    Set r = sh.Range("A1")
    
    If r.Value = "6P" Then
        tmp = True
    End If
    
    If sh.Cells(2, 1).Value Like "*TOTAL FMA*" Then
        tmp = tmp And True
    End If
    
    If sh.Range("C1").Value Like "*Y*CW*" Then
        tmp = tmp And True
    End If
    
    If sh.Range("G1").Value = "IN SCOPE" Then
        tmp = tmp And True
    End If
    
    If sh.Range("O1").Value <> "" Then
        tmp = tmp And True
    End If
    
    checkWizardBufferCells = tmp
    
End Function




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
                        MsgBox "Jest juz taki projekt - dane z nowym CW dodane na dnie tabeli."
                        Set r = r.End(xlDown).End(xlDown).End(xlUp).Offset(1, 0)
                        Exit Do
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
    
    
        If Trim(Me.TextBoxProj.Value) <> "" Then
        
    
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
        Else
            MsgBox "Brak nazwy projektu!"
        End If
    End If
End Sub

Public Sub dane_dla_nowego_skopiuj_ze_starego(old_tl As T_Link, new_tl As T_Link, m As Worksheet)

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
        For X = SIXP.e_link_cw To ostatnia_kolumna_arkusza - 1
            new_r.Offset(0, X) = old_r.Offset(0, X)
        Next X
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
