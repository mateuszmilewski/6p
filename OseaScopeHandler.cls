VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OseaScopeHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' klasa ta ma sie zajmowac kontrola danych miedzy order releasami a main ogolnie
' dobrze by bylo potrfila wpisywac ladnie aktualny ycw dla update'u arkusza
' order release, dzieki czemu bedziemy mogli miec porownanie z aktualnym cw w arkuszu main

Private m As Worksheet
Private osea As Worksheet

Private r As Range

Private Sub Class_Initialize()


    ' ten pseudo konstruktor nie bedzie mial narazie wiecej kodu
    ' 2016-08-10

    ' main sheets assigning
    ' ====================================================================
    ' przypisanie arkusza main
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' arkusz order release status handler
    Set osea = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set osea = Nothing
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormOseaScope)


    If CStr(eca) = SIXP.G_BTN_TEXT_ADD Then
    
        With o
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            
            'Public Enum E_OSEA_ORDER
            '    e_osea_order_total = 5
            '    e_osea_order_on_stock
            '    e_osea_order_ordered
            '    e_osea_order_confirmed
            '    e_osea_order_for_mrd
            '    e_osea_order_after_mrd
            '    e_osea_order_open
            'End Enum
            
            
            ' pola liczbowe
            ' ---------------------------------------------------------
            .TextBoxAfterMRD = "0"
            .TextBoxConfirmed = "0"
            .TextBoxForMRD = "0"
            .TextBoxOnStock = "0"
            .TextBoxOPEN = "0"
            .TextBoxOrdered = "0"
            .TextBoxTotal = "0"
            
            ' ---------------------------------------------------------
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            Set r = osea.Cells(1, 1)
            Do
                If CStr(.LabelTitle.Caption) = _
                    CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                

                        
                        ' pola liczbowe
                        ' ---------------------------------------------------------
                        ' new
                        '.TextBoxNoOfVeh = "0"
                        '.TextBoxWeeksDelay = "0"
                        ' EDIT
                        .TextBoxAfterMRD = CStr(r.Offset(0, SIXP.e_osea_order_after_mrd - 1))
                        .TextBoxConfirmed = CStr(r.Offset(0, SIXP.e_osea_order_confirmed - 1))
                        .TextBoxForMRD = CStr(r.Offset(0, SIXP.e_osea_order_for_mrd - 1))
                        .TextBoxOnStock = CStr(r.Offset(0, SIXP.e_osea_order_on_stock - 1))
                        .TextBoxOPEN = CStr(r.Offset(0, SIXP.e_osea_order_open - 1))
                        .TextBoxOrdered = CStr(r.Offset(0, SIXP.e_osea_order_ordered - 1))
                        .TextBoxTotal = CStr(r.Offset(0, SIXP.e_osea_order_total - 1))
                        
                        ' ---------------------------------------------------------
                    
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            
            
        
        End With
    End If
End Sub




