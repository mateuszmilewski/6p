VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XQ6PHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta ma sie zajmowac kontrola danych miedzy order releasami a main ogolnie
' dobrze by bylo potrfila wpisywac ladnie aktualny ycw dla update'u arkusza
' order release, dzieki czemu bedziemy mogli miec porownanie z aktualnym cw w arkuszu main

Private m As Worksheet
Private t As Worksheet

Private r As Range

Private Sub Class_Initialize()


    ' ten pseudo konstruktor nie bedzie mial narazie wiecej kodu
    ' 2016-08-10

    ' main sheets assigning
    ' ====================================================================
    ' przypisanie arkusza main
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' arkusz order release status handler
    Set t = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set t = Nothing
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormX6P)


    If CStr(eca) = SIXP.G_BTN_TEXT_ADD Then
    
        With o
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            
            
            .TextBoxCmnt = "your comment... "
            .DTPickerPPAPGate = Date
            .TextBoxPPAPGate = CStr(parse_from_date_to_yyyycw(Date))
            
            
            add_combobox_for_project_type o
            
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            Set r = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
            Do
                If CStr(.LabelTitle.Caption) = _
                    CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                
                        .TextBoxCmnt = CStr(r.Offset(0, SIXP.e_xq_comment - 1))
                        .TextBoxPPAPGate = CStr(r.Offset(0, SIXP.e_xq_ppap_gate - 1))
                        .DTPickerPPAPGate = SIXP.GlobalFooModule.from_yyyy_cw_to_monday_from_this_week(CStr(r.Offset(0, SIXP.e_xq_ppap_gate - 1)))
                        
                        
                        add_combobox_for_project_type o
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            
            
        
        End With
    End If
End Sub


Private Sub add_combobox_for_project_type(o As FormX6P)
        ' combo box fill
        ' ---------------
        ' ---------------
        
        o.ComboBoxXqProjectType.Clear
        
        Dim pt As Range
        Set pt = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("F2")
        
        Do
            o.ComboBoxXqProjectType.AddItem CStr(pt)
            Set pt = pt.Offset(1, 0)
        Loop Until Trim(pt) = ""
        
        o.ComboBoxXqProjectType.Value = CStr(ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("F3"))
        
        
        ' ---------------
        ' ---------------
End Sub



