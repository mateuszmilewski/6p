VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewProjectHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ac As Range

Private Sub Class_Initialize()


    If sprawdz_czy_moge_przypisac_aktywna_komorke() Then
        Set ac = ActiveCell
        
        If ac.Row = 1 Then
            Set ac = ac.Offset(1, 0)
        ElseIf ac.Row = 2 Then
        Else
            If ac <> "" Then
            Else
            
                ' to rozwiazanie jest smieszne, ale w sumie niech bedzie
                ' jesli nie znalazl w aktywnym arkuszu zadnych info
                ' to wtedy pobiera labelki i to one wklejaja sie do fieldow
                Set ac = ac.End(xlUp)
            End If
        End If
        
        With NewProj
        
            .TextBoxProj = Trim(ac.Parent.Cells(ac.Row, SIXP.e_link_project))
            .TextBoxFaza = Trim(ac.Parent.Cells(ac.Row, SIXP.e_link_faza))
            .TextBoxSelectedCW = Trim(ac.Parent.Cells(ac.Row, SIXP.e_link_cw))
            .TextBoxCW = Trim(ac.Parent.Cells(ac.Row, SIXP.e_link_cw))
            .TextBoxPlt = Trim(ac.Parent.Cells(ac.Row, SIXP.e_link_plt))
            
            
            prepare_comboboxes_and_dt_picker
            
            
            
            .Show ' vbModeless
        End With
    Else
        MsgBox "Aktywny arkussz nie jest dozwolony - przejdz do arkusza zawierajacego dane raportu."
        Set ac = Nothing
    End If
End Sub

Private Sub prepare_comboboxes_and_dt_picker()
    
    With NewProj
        .ComboBoxStatus.Clear
        
        For x = 2 To 4
        
            .ComboBoxStatus.AddItem CStr(ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Cells(x, 1))
            
            If x = 2 Then
                .ComboBoxStatus.Value = _
                    CStr(ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Cells(x, 1))
            End If
        Next x
        
        
        .ComboBoxFAZA.Clear
        
        Dim r As Range
        Set r = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("C2")
        Do
            ' ----------------------------------
            .ComboBoxFAZA.AddItem CStr(r)
            ' ----------------------------------
            Set r = r.Offset(1, 0)
        Loop Until r = ""
        
        .ComboBoxPLT.Clear
        Set r = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("D2")
        Do
            ' ----------------------------------
            .ComboBoxPLT.AddItem CStr(r)
            ' ----------------------------------
            Set r = r.Offset(1, 0)
        Loop Until r = ""
        
        
        ' no i ostatni date picker nie musi miec wypelnienia damy tylko today date
        .DTPicker1 = Date
        .TextBoxCW = SIXP.GlobalFooModule.parse_from_date_to_yyyycw(Date)
        
        
    End With
End Sub

Private Sub Class_Terminate()
    Set ac = Nothing
End Sub


Private Function sprawdz_czy_moge_przypisac_aktywna_komorke() As Boolean
    
    sprawdz_czy_moge_przypisac_aktywna_komorke = False
    
    If ActiveWorkbook.FullName = ThisWorkbook.FullName Then
        If ok_jestesmy_juz_w_srodku_prawidlowego_pliku_zatem_sprawdz_teraz_czy_jest_z_czego_zbierac_dane() Then
            sprawdz_czy_moge_przypisac_aktywna_komorke = True
        End If
    End If
End Function

Private Function ok_jestesmy_juz_w_srodku_prawidlowego_pliku_zatem_sprawdz_teraz_czy_jest_z_czego_zbierac_dane() As Boolean
    ok_jestesmy_juz_w_srodku_prawidlowego_pliku_zatem_sprawdz_teraz_czy_jest_z_czego_zbierac_dane = False
    
    
    If Cells(1, 1) = "Projekt" And _
        Cells(1, 2) = "Plant Code" And _
        Cells(1, 3) = "Faza" And _
        Cells(1, 4) = "CW" Then
        
            ok_jestesmy_juz_w_srodku_prawidlowego_pliku_zatem_sprawdz_teraz_czy_jest_z_czego_zbierac_dane = True
    End If

End Function
