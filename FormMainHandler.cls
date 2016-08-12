VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormMainHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' SPRAWDACZ!
' ================================
' ================================
Private sprawdzacz As SixPChecker
' ================================
' ================================

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

Public Sub init(Optional link_str As String)
    ' sekcja configuracyjna i inicjacyjna dla FormMain
    
    
    If sprawdzacz.sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main() Then
    
        
        
        ' wstepnie wyplnij links za pomoca arkusza main
        wypelnij_links
    
        With FormMain
            
            
            Dim l As T_Link
            
            ' to jest bardzo ciekawa konstrukcja ktora pracuje na swoim
            ' ale jednoczesnie zwraca aktualna wartosc na combobox po wypelnieniu
            ' wczesniej oczywiscie ta wartosc wyjsciowa combobox
            ' jest definiowana na podstawie aktywnej
            ' komorki :)
            Set l = wypelnij_combo_box_na_podstawie_selekcji_i_slownika(link_str)
            
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
        adjustuj_guzik SIXP.e_main_last_update_on_order_release_status, sh, main_sh, l, paleta
        
        ' recent bp ch
        Set sh = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_recent_build_plan_changes, sh, main_sh, l, paleta
        
        ' chart cont pnoc
        Set sh = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_chart_contracted_pnoc, sh, main_sh, l, paleta
        
        ' sea
        Set sh = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_osea, sh, main_sh, l, paleta
        
        ' totals
        Set sh = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_totals, sh, main_sh, l, paleta
        
        ' xq
        Set sh = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_xq, sh, main_sh, l, paleta
        
        ' del conf
        Set sh = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_del_conf, sh, main_sh, l, paleta
        
        ' open issues
        Set sh = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_open_issues, sh, main_sh, l, paleta
        
        ' resp
        Set sh = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_resp, sh, main_sh, l, paleta
    End With
    
End Sub



Private Function wypelnij_combo_box_na_podstawie_selekcji_i_slownika(Optional link_str As String) As T_Link
    
    FormMain.ComboBoxProject.Clear
    
    dodaj_najpierw_zrodla
    
    If link_str = "" Then
    
        Set link = sprobuj_dodac_wartosci_z_aktywnej_komorki()
    Else
        args = Split(link_str, ",")
        Set link = New T_Link
        link.zrob_mnie_z_argsow Trim(args(0)), Trim(args(1)), Trim(args(2)), Trim(args(3))
        wypisz_wartosc_combo_box_z_linku link
    End If
    
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
                .project = Trim(CStr(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_project)))
                .plt = Trim(CStr(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_plt)))
                .faza = Trim(CStr(ActiveCell.Parent.Cells(ActiveCell.Row, SIXP.e_link_faza)))
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
    
    
    
    ' to jest bardzo prosty konstruktor przypisujacy tylko i wylacznie odpowiednia palete barw do zmiennych co by bylo w uzycia przez cale
    ' zycie obiektu w ktorym wlasnie sie znajdujemy
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
            link.project = Trim(r.Parent.Cells(r.Row, SIXP.e_link_project))
            link.plt = Trim(r.Parent.Cells(r.Row, SIXP.e_link_plt))
            link.faza = Trim(r.Parent.Cells(r.Row, SIXP.e_link_faza))
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


    ' ten sub zajmuje sie wlasciwie tylko sprawdzeniem
    ' czy dany arkusz w jakim sie znajdujemy (i z ktorego bedziemy chcieli sciagnac dane to te faktyczne)
    ' wszystko miesci sie w konstruktorze zatem konstrukcja jest skrajnie mala
    Dim n As NewProjectHandler
    Set n = New NewProjectHandler
    Set n = Nothing
End Sub


