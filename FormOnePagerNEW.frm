VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOnePagerNEW 
   Caption         =   "Raport Na Podstawie Selekcji w MAIN"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11370
   OleObjectBlob   =   "FormOnePagerNEW.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOnePagerNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnNewTable_Click()
    Hide
    skompletuj_dane_pod_generowania_kolejnych_one_pagerow2 E_ONE_PAGERS_INTO_NEW_TABLE, E_NEW_ONE_PAGER_LAYOUT
End Sub

Private Sub BtnPowerPoint_Click()



    Hide
    
    If Me.RadioExcels Then
    
        skompletuj_dane_pod_generowania_kolejnych_one_pagerow2 E_ONE_PAGERS_INTO_SEPERATE_EXCELS, E_NEW_ONE_PAGER_LAYOUT
    ElseIf Me.RadioPowerPoint Then
    
        skompletuj_dane_pod_generowania_kolejnych_one_pagerow2 E_ONE_PAGERS_INTO_POWER_POINT, E_NEW_ONE_PAGER_LAYOUT
    Else
    
        MsgBox "nie ma innej mozliwosci! blad krytyczny, makro zatrzymalo sie!"
        End
    End If

End Sub


Private Sub skompletuj_dane_pod_generowania_kolejnych_one_pagerow2(e As E_ONE_PAGERS_INTO, Optional eOnePagerLayout As E_ONE_PAGER_LAYOUT)


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
        If Not r.EntireRow.Hidden Then
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
