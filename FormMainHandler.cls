VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormMainHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sprawdzacz As SixPChecker
Private links As Dictionary
Private link As T_Link
Private tmp_link As T_Link

Private main_sh As Worksheet

Private sh As Worksheet

Private r As Range


Private linker_obj As Linker


' paleta pod formularz "the daily commute"
Private paleta As PaletaTheDailyCommute

Private delegacja_pod_nowy_projekt As NewProjectHandler

Public Sub init()
    ' sekcja configuracyjna i inicjacyjna dla FormMain
    
    
    If sprawdzacz.sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main() Then
    
        
        
        ' wstepnie wyplnij links za pomoca arkusza main
        wypelnij_links
    
        With FormMain
            
            
            Dim l As T_Link
            Set l = wypelnij_combo_box_na_podstawie_selekcji_i_slownika()
            
            ' wartosc w combo box okreslone
            ' czas na texty w guzikach
            adjust_na_guzikach l
            
            ' na samym koncu gdy wszystko mamy show
            .Show ' vbModeless
        End With
    Else
        MsgBox "niestety z tego poziomu (arkusza) iterfejs nie moze zostac otwarty"
        End
    End If
End Sub

Private Sub adjust_na_guzikach(ByRef l As T_Link)
    
    ' slownik links dalej jest aktywny
    With FormMain
        
        ' order release status
        Set sh = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_order_release_status, sh, main_sh, l
        
        ' recent bp ch
        Set sh = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_recent_build_plan_changes, sh, main_sh, l
        
        ' chart cont pnoc
        Set sh = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_chart_contracted_pnoc, sh, main_sh, l
        
        ' sea
        Set sh = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_osea, sh, main_sh, l
        
        ' totals
        Set sh = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_totals, sh, main_sh, l
        
        ' xq
        Set sh = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_xq, sh, main_sh, l
        
        ' del conf
        Set sh = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_del_conf, sh, main_sh, l
        
        ' open issues
        Set sh = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_open_issues, sh, main_sh, l
        
        ' resp
        Set sh = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_resp, sh, main_sh, l
    End With
    
End Sub

Public Sub adjustuj_guzik(e_main As E_MAIN_ORDER, ish As Worksheet, imainsh As Worksheet, il As T_Link)
    
    
    Dim range_from_main As Range, range_from_ish As Range
    
    ' porownanie posrednie
    Set range_from_main = il.znajdz_siebie_w_arkuszu(imainsh)
    Set range_from_ish = il.znajdz_siebie_w_arkuszu(ish)
    
    ' jesli te dwie zmienne nie sa puste to znaczy ze mamy takie dane spasowane i chcemy je edytowac
    If Not range_from_main Is Nothing Then
        If Not range_from_ish Is Nothing Then
            
            If e_main = e_main_last_update_on_order_release_status Then
            
                With FormMain.BtnOrderReleaseStatus
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_recent_build_plan_changes Then
                
                With FormMain.BtnRecentBuildPlanChanges
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_chart_contracted_pnoc Then
                
                With FormMain.BtnContractedPNOC
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_osea Then
                
                With FormMain.BtnOseaScope
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_totals Then
                
                With FormMain.BtnTotals
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_xq Then
                
                With FormMain.BtnXq
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_del_conf Then
                
                With FormMain.BtnDelConf
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_open_issues Then
                
                With FormMain.BtnOpenIssues
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_resp Then
                
                With FormMain.BtnResp
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            End If
        Else
            ' sekcja, gdzie cos znalezlismy w main jednak nie ma tego w arkuszu przeszukiwanym
            ' to znaczy tyle ze trzeba dodac nowe
            
            If e_main = e_main_last_update_on_order_release_status Then
                FormMain.BtnOrderReleaseStatus.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOrderReleaseStatus.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_recent_build_plan_changes Then
                FormMain.BtnRecentBuildPlanChanges.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnRecentBuildPlanChanges.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_chart_contracted_pnoc Then
                FormMain.BtnContractedPNOC.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnContractedPNOC.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_osea Then
                FormMain.BtnOseaScope.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOseaScope.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_totals Then
                FormMain.BtnTotals.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnTotals.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_xq Then
                FormMain.BtnXq.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnXq.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_del_conf Then
                FormMain.BtnDelConf.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnDelConf.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_open_issues Then
                FormMain.BtnOpenIssues.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOpenIssues.BackColor = paleta.yellow
            ElseIf e_main = e_main_last_update_on_resp Then
                FormMain.BtnResp.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnResp.BackColor = paleta.yellow
            End If
        End If
    End If
    
    
    
    
End Sub

Private Function wypelnij_combo_box_na_podstawie_selekcji_i_slownika() As T_Link
    
    FormMain.ComboBoxProject.Clear
    
    dodaj_najpierw_zrodla
    
    Set link = sprobuj_dodac_wartosci_z_aktywnej_komorki()
    
    Set wypelnij_combo_box_na_podstawie_selekcji_i_slownika = link
    
End Function


Private Sub dodaj_najpierw_zrodla()
    
    For Each klucz In links.Keys
    
        With FormMain
            
            .ComboBoxProject.AddItem links(klucz).project & ", " & links(klucz).plt & ", " & links(klucz).faza & ", " & CStr(links(klucz).cw)
        End With
    Next
End Sub

Private Function sprobuj_dodac_wartosci_z_aktywnej_komorki() As T_Link


    Set sprobuj_dodac_wartosci_z_aktywnej_komorki = Nothing
    
    If ActiveCell.Row > 1 Then
        With FormMain
        
        
            Set link = New T_Link
            With link
                .project = CStr(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_project))
                .plt = CStr(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_plt))
                .faza = CStr(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_faza))
                On Error Resume Next
                .cw = CLng(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_cw))
            End With
            
            match_z_aktywna_komorka_i_zrodlem_main_sh = False
            For Each klucz In links.Keys
            
            
                Set tmp_link = links(klucz)
                
                If link.porownaj_z_drugim_linkiem(tmp_link) Then
                    match_z_aktywna_komorka_i_zrodlem_main_sh = True
                End If
            Next
            
            If match_z_aktywna_komorka_i_zrodlem_main_sh Then
                wypisz_wartosc_combo_box_z_linku link
                Set sprobuj_dodac_wartosci_z_aktywnej_komorki = link
            Else
                wypisz_wartosc_combo_box_z_linku tmp_link
                Set sprobuj_dodac_wartosci_z_aktywnej_komorki = tmp_link
            End If
            
            
            
        End With
    Else
        MsgBox "komorka pierwszego wiersza w arkuszu main nie moze byc aktywna podczas uruchamiania formularza!"
        End
    End If
End Function

Private Sub wypisz_wartosc_combo_box_z_linku(l As T_Link)
    
    With FormMain
        .ComboBoxProject.Value = l.project & ", " & l.plt & ", " & l.faza & ", " & l.cw
    End With
End Sub

Private Sub Class_Initialize()

    Set sprawdzacz = New SixPChecker
    
    Set main_sh = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    Set link = Nothing
    Set linker_obj = New Linker
    
    Set paleta = New PaletaTheDailyCommute

End Sub

Private Sub Class_Terminate()
    Set sprawdzacz = Nothing
    Set links = Nothing
    Set link = Nothing
    Set linker_obj = Nothing
    Set paleta = Nothing
End Sub



Private Sub wypelnij_links()
    
    Set links = New Dictionary
    Set r = main_sh.Range("A2")
    Do
        If Trim(r) <> "" Then
            ' sa dane to lecimy z dodawaniem do slownika
            ' -------------------------------------------
            Set link = New T_Link
            link.project = r.Parent.Cells(r.Row, SIXP.e_link_project)
            link.plt = r.Parent.Cells(r.Row, SIXP.e_link_plt)
            link.faza = r.Parent.Cells(r.Row, SIXP.e_link_faza)
            On Error Resume Next
            link.cw = CLng(r.Parent.Cells(r.Row, SIXP.e_link_cw))
            
            ajdik = linker_obj.retrun_short_number(link)
            
            If Not links.Exists(ajdik) Then
                links.Add ajdik, link
            End If
            
            
            ' -------------------------------------------
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub


Public Sub new_project()
    Dim n As NewProjectHandler
    Set n = New NewProjectHandler
    Set n = Nothing
End Sub


