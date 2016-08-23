Attribute VB_Name = "ClearModule"
Public Sub clear_item(ictrl As IRibbonControl)
    innerclearitem
End Sub


Private Sub innerclearitem()
    
    Dim s As SixPChecker
    Set s = New SixPChecker
    
    If s.sprawdz_czy_arkusz_aktywny_to_ten_arkusz Then
        ' to jest ogolenie sprawdzenie czy pierwsze 4 kolumny to kolumny std proj, plt, faza, cw
        If s.sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main Then
            
            Dim dm As DeletionManager
            Set dm = New DeletionManager
            
            
            dm.usun_kazde_wystapienie_dla_aktywnej_komorki ActiveCell
            
            Set dm = Nothing
        End If
    End If
End Sub
