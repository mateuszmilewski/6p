VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOnePager 
   Caption         =   "One Pagers Generator"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11235
   OleObjectBlob   =   "FormOnePager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOnePager"
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

Public czy_uruchamiamy_eventy As Boolean

Private Sub BtnReset_Click()
    
    SIXP.GlobalFooModule.gotoThisWorkbookMainA1
    
    Hide
    
    ' clear_one_pager
    skonfiguruj_form_generowania_one_pagera
    FormOnePager.Show vbModeless
End Sub

Private Sub BtnSubmit_Click()
    
    Hide
    
    If Me.RadioExcels Then
    
        skompletuj_dane_pod_generowania_kolejnych_one_pagerow E_ONE_PAGERS_INTO_SEPERATE_EXCELS, E_OLD_ONE_PAGER_LAYOUT
    ElseIf Me.RadioPowerPoint Then
    
        skompletuj_dane_pod_generowania_kolejnych_one_pagerow E_ONE_PAGERS_INTO_POWER_POINT, E_OLD_ONE_PAGER_LAYOUT
    Else
    
        MsgBox "nie ma innej mozliwosci! blad krytyczny, makro zatrzymalo sie!"
        End
    End If
End Sub

Private Sub skompletuj_dane_pod_generowania_kolejnych_one_pagerow(e As E_ONE_PAGERS_INTO, Optional eOnePagerLayout As E_ONE_PAGER_LAYOUT)


    Application.ScreenUpdating = False
    
    SIXP.LoadingFormModule.showLoadingForm
    
    Dim kolekcja_linkow As Collection
    Set kolekcja_linkow = New Collection
    
    Dim l As T_Link
    
    Dim m As Worksheet
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    Dim r As Range
    Set r = m.Cells(2, 1)
    
    Do
        If czy_ten_wiersz_jest_dopasowany_do_selekcji(r) Then
            Set l = New T_Link
            l.zrob_mnie_z_range r
            kolekcja_linkow.Add l
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    
    
    ' heurystycznie dobieram zbior ilosci raportow miedzy o a 100 wydaje sie rozsadnym by nie pozwalac generowac wiecej niz 100 raportow czyz nie? :D
    If kolekcja_linkow.Count > 0 And kolekcja_linkow.Count < 100 Then
    
        
        ans = MsgBox("Logika wygenerowala: " & kolekcja_linkow.Count & " potencjalnych raportow; chcesz kontynuowac?", vbYesNo)
        
        
        If ans = vbYes Then
        
        
            If eOnePagerLayout = E_NEW_ONE_PAGER_LAYOUT Then
                
                ' zupelnie nowa logika
                ' ---------------------------------------------------------------------------
                Dim noph As NewOnePagerHandler
                Set noph = New NewOnePagerHandler
                
                noph.przypisz_kolekcje_linkow kolekcja_linkow
                noph.generuj_raporty e
                
                Set noph = Nothing
                
                ' ---------------------------------------------------------------------------
            Else
            
                ' old layout
                '' ---------------------------------------------------------------------------
                
                Dim oph As OnePagerHandler
                Set oph = New OnePagerHandler
                
                oph.przypisz_kolekcje_linkow kolekcja_linkow
                oph.generuj_raporty e
                
                Set oph = Nothing
                
                '' ---------------------------------------------------------------------------
            End If
            
        Else
            MsgBox "Nic nie zostanie wygenerowane... Dobranoc :D!"
        End If
    Else
        MsgBox "niewlasciwa konfiguracja startowa... err: !(kolekcja_linkow.Count > 0 And kolekcja_linkow.Count < 100)"
    End If
    
    Set kolekcja_linkow = Nothing
    
    SIXP.LoadingFormModule.hideLoadingForm
    
    Application.ScreenUpdating = True
    
    
    MsgBox "ready!"
End Sub

Private Function czy_ten_wiersz_jest_dopasowany_do_selekcji(r As Range) As Boolean
    czy_ten_wiersz_jest_dopasowany_do_selekcji = False
    
    dopasowanie_projektu = False
    dopasowanie_plantu = False
    dopasowanie_fazy = False
    dopasowanie_cw = False
    

    For x = 0 To Me.ListBoxProjects.ListCount - 1

        
        If Me.ListBoxProjects.List(x) = Trim(r) Then
            dopasowanie_projektu = True
        End If

    Next x
    

    For x = 0 To Me.ListBoxPlants.ListCount - 1
        
        If Me.ListBoxPlants.List(x) = Trim(r.Offset(0, 1)) Then
            dopasowanie_plantu = True
        End If

    Next x
    
    For x = 0 To Me.ListBoxPhases.ListCount - 1

    
        If Me.ListBoxPhases.List(x) = Trim(r.Offset(0, 2)) Then
            dopasowanie_fazy = True
        End If

    Next x

    For x = 0 To Me.ListBoxCWs.ListCount - 1

        
        If Me.ListBoxCWs.List(x) = Trim(r.Offset(0, 3)) Then
            dopasowanie_cw = True
        End If

    Next x
    
    
    czy_ten_wiersz_jest_dopasowany_do_selekcji = _
        dopasowanie_projektu And _
        dopasowanie_plantu And _
        dopasowanie_fazy And _
        dopasowanie_cw
End Function

Private Sub BtnSubmitNew_Click()
    Hide
    
    If Me.RadioExcels Then
    
        skompletuj_dane_pod_generowania_kolejnych_one_pagerow E_ONE_PAGERS_INTO_SEPERATE_EXCELS, E_NEW_ONE_PAGER_LAYOUT
    ElseIf Me.RadioPowerPoint Then
    
        skompletuj_dane_pod_generowania_kolejnych_one_pagerow E_ONE_PAGERS_INTO_POWER_POINT, E_NEW_ONE_PAGER_LAYOUT
    Else
    
        MsgBox "nie ma innej mozliwosci! blad krytyczny, makro zatrzymalo sie!"
        End
    End If
End Sub

Private Sub ListBoxCWs_Change()
    inner_change_on_listboxes
End Sub

Private Sub ListBoxPhases_Change()
    inner_change_on_listboxes
End Sub

Private Sub ListBoxPlants_Change()

    inner_change_on_listboxes
End Sub


Private Sub ListBoxProjects_Change()
    inner_change_on_listboxes
End Sub

Private Sub inner_change_on_listboxes()


    If Me.czy_uruchamiamy_eventy Then


        Application.EnableEvents = False
        
        Dim proj_dic As Dictionary
        Dim plt_dic As Dictionary
        Dim faza_dic As Dictionary
        Dim cw_dic As Dictionary
        
        
        Set proj_dic = New Dictionary
        Set plt_dic = New Dictionary
        Set faza_dic = New Dictionary
        Set cw_dic = New Dictionary
    
    
    
        pobierz_wszystkie_selekcje proj_dic, plt_dic, faza_dic, cw_dic
        przejrzyj_na_podstawie_selekcji_jeszcze_raz_zbior_danych proj_dic, plt_dic, faza_dic, cw_dic
        przeorganizuj_listboxy proj_dic, plt_dic, faza_dic, cw_dic
        
        Set proj_dic = Nothing
        Set plt_dic = Nothing
        Set faza_dic = Nothing
        Set cw_dic = Nothing
        
        Application.EnableEvents = True
    End If
End Sub

Private Sub przeorganizuj_listboxy( _
    ByRef proj_dic As Dictionary, _
    ByRef plt_dic As Dictionary, _
    ByRef faza_dic As Dictionary, _
    ByRef cw_dic As Dictionary)
    
    
    With Me
    
        .ListBoxProjects.Clear
        .ListBoxPlants.Clear
        .ListBoxPhases.Clear
        .ListBoxCWs.Clear
        
    End With
    
    
    x = 0
    For Each Key In proj_dic.Keys
        
        If czy_key_nie_byl_jeszcze_wsadzony(CStr(Key), SIXP.e_link_project) Then
            Me.ListBoxProjects.AddItem CStr(Key)
            
            If proj_dic(Key) = 1 Then
                czy_uruchamiamy_eventy = False
                Me.ListBoxProjects.Selected(x) = True
                czy_uruchamiamy_eventy = True
            End If
            x = x + 1
        End If
    Next
    
    x = 0
    For Each Key In plt_dic.Keys
        
        If czy_key_nie_byl_jeszcze_wsadzony(CStr(Key), SIXP.e_link_plt) Then
            Me.ListBoxPlants.AddItem CStr(Key)
        
            If plt_dic(Key) = 1 Then
                Application.EnableEvents = False
                czy_uruchamiamy_eventy = False
                Me.ListBoxPlants.Selected(x) = True
                czy_uruchamiamy_eventy = True
                Application.EnableEvents = True
            End If
            x = x + 1
        End If
    Next
    
    x = 0
    For Each Key In faza_dic.Keys
        If czy_key_nie_byl_jeszcze_wsadzony(CStr(Key), SIXP.e_link_faza) Then
            Me.ListBoxPhases.AddItem CStr(Key)
            
            If faza_dic(Key) = 1 Then
                czy_uruchamiamy_eventy = False
                Me.ListBoxPhases.Selected(x) = True
                czy_uruchamiamy_eventy = True
            End If
            x = x + 1
        End If
    Next
    
    x = 0
    For Each Key In cw_dic.Keys
        
        If czy_key_nie_byl_jeszcze_wsadzony(CStr(Key), SIXP.e_link_cw) Then
            Me.ListBoxCWs.AddItem CStr(Key)
            
            If cw_dic(Key) = 1 Then
                czy_uruchamiamy_eventy = False
                Me.ListBoxCWs.Selected(x) = True
                czy_uruchamiamy_eventy = True
            End If
            x = x + 1
        End If
    Next

End Sub


Private Function czy_key_nie_byl_jeszcze_wsadzony(k As String, e As E_LINK_ORDER) As Boolean
    czy_key_nie_byl_jeszcze_wsadzony = True
    
    
    
    If e = e_link_project Then
        For x = 0 To Me.ListBoxProjects.ListCount - 1
            If CStr(Me.ListBoxProjects.List(x)) = CStr(k) Then
                czy_key_nie_byl_jeszcze_wsadzony = False
                Exit Function
            End If
        Next x
        
        
    ElseIf e = e_link_plt Then
        For x = 0 To Me.ListBoxPlants.ListCount - 1
            If CStr(Me.ListBoxPlants.List(x)) = CStr(k) Then
                czy_key_nie_byl_jeszcze_wsadzony = False
                Exit Function
            End If
        Next x
        
        
    ElseIf e = e_link_faza Then
        For x = 0 To Me.ListBoxPhases.ListCount - 1
            If CStr(Me.ListBoxPhases.List(x)) = CStr(k) Then
                czy_key_nie_byl_jeszcze_wsadzony = False
                Exit Function
            End If
        Next x
        
        
    ElseIf e = e_link_cw Then
        For x = 0 To Me.ListBoxCWs.ListCount - 1
            If CStr(Me.ListBoxCWs.List(x)) = CStr(k) Then
                czy_key_nie_byl_jeszcze_wsadzony = False
                Exit Function
            End If
        Next x


    End If
    
    
    
End Function


Private Sub pobierz_wszystkie_selekcje( _
    ByRef proj_dic As Dictionary, _
    ByRef plt_dic As Dictionary, _
    ByRef faza_dic As Dictionary, _
    ByRef cw_dic As Dictionary)
    
    ' w pierwszej kolejnosci musimy sprawdzic jakie selekcje zostaly wykonane przez usera
    ' rozwiazanie musi byc klarowne i zawierac krzyzowe dopasowanie danych - niestety w najgorszym razie jest to
    ' poczworna petla ...
    
    For x = 0 To Me.ListBoxProjects.ListCount - 1
        If Me.ListBoxProjects.Selected(x) Then
            If Not proj_dic.Exists(Me.ListBoxProjects.List(x)) Then
                proj_dic.Add CStr(Me.ListBoxProjects.List(x)), 1
            End If
        End If
    Next x
    
    For x = 0 To Me.ListBoxPlants.ListCount - 1
        If Me.ListBoxPlants.Selected(x) Then
            If Not plt_dic.Exists(Me.ListBoxPlants.List(x)) Then
                plt_dic.Add CStr(Me.ListBoxPlants.List(x)), 1
            End If
        End If
    Next x
    
    For x = 0 To Me.ListBoxPhases.ListCount - 1
        If Me.ListBoxPhases.Selected(x) Then
            If Not faza_dic.Exists(Me.ListBoxPhases.List(x)) Then
                faza_dic.Add CStr(Me.ListBoxPhases.List(x)), 1
            End If
        End If
    Next x
    
    For x = 0 To Me.ListBoxCWs.ListCount - 1
        If Me.ListBoxCWs.Selected(x) Then
            If Not cw_dic.Exists(Me.ListBoxCWs.List(x)) Then
                cw_dic.Add CStr(Me.ListBoxCWs.List(x)), 1
            End If
        End If
    Next x
End Sub

Private Sub przejrzyj_na_podstawie_selekcji_jeszcze_raz_zbior_danych( _
    ByRef proj_dic As Dictionary, _
    ByRef plt_dic As Dictionary, _
    ByRef faza_dic As Dictionary, _
    ByRef cw_dic As Dictionary)
    
    ' mamy juz slwoniki wypelnione - teraz trzeba je przejrzec
    
    
    
    Dim m As Worksheet
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    Dim r As Range
    Set r = m.Cells(2, 1)
    
    'For Each Key In proj_dic.Keys
    '    MsgBox Key
    'Next
    
    ' should be true
    ' bbbol = proj_dic.Exists(CStr(r.Offset(0, SIXP.e_link_project - 1)))
    
    
        
    
    
    Dim str_proj_key As String
    Dim str_plt_key As String
    Dim str_faza_key As String
    Dim str_cw_key As String
    
    
    Do
    
        
        
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        str_z_maina = CStr(r.Offset(0, SIXP.e_link_project - 1)) & _
            SIXP.G_SEPARATOR & CStr(r.Offset(0, SIXP.e_link_plt - 1)) & _
            SIXP.G_SEPARATOR & CStr(r.Offset(0, SIXP.e_link_faza - 1)) & _
            SIXP.G_SEPARATOR & CStr(r.Offset(0, SIXP.e_link_cw - 1))
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
            
            
            
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        If proj_dic.Count = 0 Then
            str_proj_key = "*"
        ElseIf proj_dic.Exists(CStr(r.Offset(0, SIXP.e_link_project - 1))) Then
        
            If proj_dic(CStr(r.Offset(0, SIXP.e_link_project - 1))) = 1 Then
                str_proj_key = CStr(r.Offset(0, SIXP.e_link_project - 1))
            Else
                str_proj_key = "*"
            End If
        Else
            str_proj_key = CStr(wrzuc_not_found_lub_gwiazdke(proj_dic))
        End If
        
        If plt_dic.Count = 0 Then
            str_plt_key = "*"
        ElseIf plt_dic.Exists(CStr(r.Offset(0, SIXP.e_link_plt - 1))) Then
        
            If plt_dic(CStr(r.Offset(0, SIXP.e_link_plt - 1))) = 1 Then
                str_plt_key = CStr(r.Offset(0, SIXP.e_link_plt - 1))
            Else
                str_plt_key = "*"
            End If
        Else
            str_plt_key = CStr(wrzuc_not_found_lub_gwiazdke(plt_dic))
        End If
        
        If faza_dic.Count = 0 Then
            str_faza_key = "*"
        ElseIf faza_dic.Exists(CStr(r.Offset(0, SIXP.e_link_faza - 1))) Then
            
            If faza_dic(CStr(r.Offset(0, SIXP.e_link_faza - 1))) = 1 Then
                str_faza_key = CStr(r.Offset(0, SIXP.e_link_faza - 1))
            Else
                str_faza_key = "*"
            End If
        Else
            str_faza_key = CStr(wrzuc_not_found_lub_gwiazdke(faza_dic))
        End If
        
        If cw_dic.Count = 0 Then
            str_cw_key = "*"
        ElseIf cw_dic.Exists(CStr(r.Offset(0, SIXP.e_link_cw - 1))) Then
            If cw_dic(CStr(r.Offset(0, SIXP.e_link_cw - 1))) = 1 Then
                str_cw_key = CStr(r.Offset(0, SIXP.e_link_cw - 1))
            Else
                str_cw_key = "*"
            End If
        Else
            str_cw_key = CStr(wrzuc_not_found_lub_gwiazdke(cw_dic))
        End If
            
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
    
    
        ' --------------------------------------------------------------' --------------------------------------------------------------
        str_z_dicsow = str_proj_key & SIXP.G_SEPARATOR & str_plt_key & SIXP.G_SEPARATOR & str_faza_key & SIXP.G_SEPARATOR & str_cw_key
        ' --------------------------------------------------------------' --------------------------------------------------------------
    
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        If CStr(str_z_maina) Like CStr(str_z_dicsow) Then
            'Dim tprojd As Dictionary
            'Dim tpltd As Dictionary
            'Dim tfd As Dictionary
            'Dim tcwd As Dictionary
            dodanie_elementu r, proj_dic, plt_dic, faza_dic, cw_dic
        End If
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        ' --------------------------------------------------------------
        
        
        Set r = r.Offset(1, 0)
        
    Loop Until Trim(r) = ""
    
    
End Sub

Private Function wrzuc_not_found_lub_gwiazdke(ByRef d As Dictionary) As String
    
    wrzuc_not_found_lub_gwiazdke = "NOT FOUND"
    
    
    For Each k In d.Keys
        If d(CStr(k)) = 1 Then
            wrzuc_not_found_lub_gwiazdke = CStr(k)
            Exit Function
        Else
            wrzuc_not_found_lub_gwiazdke = "*"
        End If
    Next
    
    
End Function


    

Private Sub dodanie_elementu(ByRef r As Range, _
    ByRef proj_dic As Dictionary, _
    ByRef plt_dic As Dictionary, _
    ByRef faza_dic As Dictionary, _
    ByRef cw_dic As Dictionary)
    
    
    ' r = link proj
    
    If Not proj_dic.Exists(CStr(r)) Then
        proj_dic.Add CStr(r), 0
    End If
    
    If Not plt_dic.Exists(CStr(r.Offset(0, SIXP.e_link_plt - 1))) Then
        plt_dic.Add CStr(r.Offset(0, SIXP.e_link_plt - 1)), 0
    End If
    
    If Not faza_dic.Exists(CStr(r.Offset(0, SIXP.e_link_faza - 1))) Then
        faza_dic.Add CStr(r.Offset(0, SIXP.e_link_faza - 1)), 0
    End If
    
    If Not cw_dic.Exists(CStr(r.Offset(0, SIXP.e_link_cw - 1))) Then
        cw_dic.Add CStr(r.Offset(0, SIXP.e_link_cw - 1)), 0
    End If
    
End Sub
    



Private Function sprawdz_czy_jest_zaznaczona_wartosc(e As E_LINK_ORDER) As Boolean
    
    sprawdz_czy_jest_zaznaczona_wartosc = False
    
    
    Dim lb As Variant
    
    If e = e_link_project Then
        Set lb = FormOnePager.ListBoxProjects
    ElseIf e = e_link_plt Then
        Set lb = FormOnePager.ListBoxPlants
    ElseIf e = e_link_faza Then
        Set lb = FormOnePager.ListBoxPhases
    ElseIf e = e_link_cw Then
        Set lb = FormOnePager.ListBoxCWs
    Else
        Set lb = Nothing
    End If
    
    If Not lb Is Nothing Then
    
        For x = 0 To lb.ListCount - 1
            If lb.Selected(x) = True Then
                
                sprawdz_czy_jest_zaznaczona_wartosc = True
                Exit Function
            End If
        
        Next x
        
    End If
End Function
