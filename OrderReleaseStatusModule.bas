Attribute VB_Name = "OrderReleaseStatusModule"
' sekcja subow, ktore pracuja po kliknieciu w odpowiednie guziki
Public Sub zrob_order_release_status(edit_czy_add As String, txt_from_combo_box As String)

    Dim o As OrderReleaseStatusHandler
    Set o = New OrderReleaseStatusHandler
    
    
    With FormOrderReleaseStatus
        .LabelTitle.Caption = CStr(txt_from_combo_box)
        .BtnSubmit.Caption = CStr(edit_czy_add)
        
        .Show
    End With
    
    Set o = Nothing

End Sub
