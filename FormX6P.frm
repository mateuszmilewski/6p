VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormX6p 
   Caption         =   "X6p"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormX6P.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormX6P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnSubmit_Click()

    ' text na guziki
    ' Global Const G_BTN_TEXT_ADD = "Dodaj"
    ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
    Hide
    inner_calc
    
    run_FormMain Me.LabelTitle
End Sub

Private Sub change_col_K_in_MAIN_worksheet(ByRef r As Range)
    
    ' tutaj sekcja, gdy dane juz zostaly dodane do arkusza order releases
    ' teraz nalezy odpowiednio o tym poinformowac arkusz glowny
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
    
        ' szukamy teraz w main
        ' ===================================================
        Dim rr As Range
        Set rr = ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Cells(1, 1)
        Do
            If CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(CStr(r.Offset(0, 3)))) = _
                CStr(Trim(rr) & ", " & Trim(rr.Offset(0, 1)) & ", " & Trim(rr.Offset(0, 2)) & ", " & Trim(CStr(rr.Offset(0, 3)))) Then
                    ' jest to samo w order release status sheet oraz to samo w main sheet
                    ' --------------------------------------------------------------------
                    ''
                    '
                    rr.Offset(0, SIXP.e_main_last_update_on_xq - 1) = Trim(CStr(rr.Offset(0, 3)))
                    '
                    ''
                    ' --------------------------------------------------------------------
                    Exit Do
            End If
            Set rr = rr.Offset(1, 0)
        Loop Until Trim(rr) = ""
        
        
        ' ===================================================
    
    
    
    
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
End Sub

Private Sub inner_calc()


    'Public Enum E_XQ_ORDER
    '    e_xq_comment = 5
    '    e_xq_ppap_gate
    '    e_xq_project_type
    'End Enum


    Dim r As Range
    
    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = arr(x)
        Next x
        
        
        give_data_to_ranges r
        change_col_K_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm).Cells(1, 1)
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_K_in_MAIN_worksheet r
                    Exit Do
            End If
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        ' ===================================================
    Else
        MsgBox "fatal error on submitting!"
        End
    End If
End Sub

Private Sub give_data_to_ranges(ByRef r As Range)
    r.Parent.Cells(r.Row, SIXP.e_xq_comment) = CStr(Me.TextBoxCmnt)
    r.Parent.Cells(r.Row, SIXP.e_xq_ppap_gate) = CStr(Me.TextBoxPPAPGate)
    r.Parent.Cells(r.Row, SIXP.e_xq_project_type) = CStr(Me.ComboBoxXqProjectType)
End Sub




Private Sub DTPickerPPAPGate_Change()

    ' Me.TextBoxReleased = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerOrdersDue)))
    Me.TextBoxPPAPGate = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerPPAPGate)))
End Sub
