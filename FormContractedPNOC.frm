VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormContractedPNOC 
   Caption         =   "FormContractedPNOC"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormContractedPNOC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormContractedPNOC"
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

Private Sub change_col_H_in_MAIN_worksheet(ByRef r As Range)
    
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
                    rr.Offset(0, SIXP.e_main_last_update_on_chart_contracted_pnoc - 1) = Trim(CStr(rr.Offset(0, 3)))
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


    'Public Enum E_ORDER_RELEASE_STATUS
    '    e_order_release_mrd = 5
    '    e_order_release_build
    '    e_order_release_bom_freeze
    '    e_order_release_no_of_veh
    '    e_order_release_orders_due
    '    e_order_release_released
    '    e_order_release_weeks_delay
    'End Enum


    Dim r As Range
    
    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = arr(x)
        Next x
        
        
        give_data_to_ranges r
        change_col_H_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm).Cells(1, 1)
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_H_in_MAIN_worksheet r
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
    r.Parent.Cells(r.Row, SIXP.e_cont_pnoc_chart_actionable_fma) = CStr(Me.TextBoxActionableFMA)
    r.Parent.Cells(r.Row, SIXP.e_cont_pnoc_chart_contracted) = CStr(Me.TextBoxContracted)
    r.Parent.Cells(r.Row, SIXP.e_cont_pnoc_chart_open_bp) = CStr(Me.TextBoxOpenBP)
    r.Parent.Cells(r.Row, SIXP.e_cont_pnoc_chart_pnoc) = CStr(Me.TextBoxPNOC)
End Sub





' textboxes with qtyies
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------

'Private Sub NoOfVehLess_Click()
'    If IsNumeric(Me.TextBoxNoOfVeh) Then
'        If CLng(Me.TextBoxNoOfVeh) > 0 Then
'            tmp = CLng(Me.TextBoxNoOfVeh)
'            tmp = tmp - 1
'            Me.TextBoxNoOfVeh = CStr(tmp)
'        End If
'    End If
'End Sub

'Private Sub NoOfVehMore_Click()
'    If IsNumeric(Me.TextBoxNoOfVeh) Then
'        tmp = CLng(Me.TextBoxNoOfVeh)
'        tmp = tmp + 1
'        Me.TextBoxNoOfVeh = CStr(tmp)
'    End If
'End Sub



Private Sub ActionableFMALess_Click()
    If IsNumeric(Me.TextBoxActionableFMA) Then
        If CLng(Me.TextBoxActionableFMA) > 0 Then
            tmp = CLng(Me.TextBoxActionableFMA)
            tmp = tmp - 1
            Me.TextBoxActionableFMA = CStr(tmp)
        End If
    End If
End Sub

Private Sub ActionableFMALess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxActionableFMA) Then
        If CLng(Me.TextBoxActionableFMA) > 9 Then
            tmp = CLng(Me.TextBoxActionableFMA)
            tmp = tmp - 10
            Me.TextBoxActionableFMA = CStr(tmp)
        End If
    End If
End Sub

Private Sub ActionableFMAMore_Click()
    If IsNumeric(Me.TextBoxActionableFMA) Then
        tmp = CLng(Me.TextBoxActionableFMA)
        tmp = tmp + 1
        Me.TextBoxActionableFMA = CStr(tmp)
    End If
End Sub

Private Sub ActionableFMAMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxActionableFMA) Then
        tmp = CLng(Me.TextBoxActionableFMA)
        tmp = tmp + 10
        Me.TextBoxActionableFMA = CStr(tmp)
    End If
End Sub

Private Sub ContLess_Click()
    If IsNumeric(Me.TextBoxContracted) Then
        If CLng(Me.TextBoxContracted) > 0 Then
            tmp = CLng(Me.TextBoxContracted)
            tmp = tmp - 1
            Me.TextBoxContracted = CStr(tmp)
        End If
    End If
End Sub



Private Sub ContLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxContracted) Then
        If CLng(Me.TextBoxContracted) > 9 Then
            tmp = CLng(Me.TextBoxContracted)
            tmp = tmp - 10
            Me.TextBoxContracted = CStr(tmp)
        End If
    End If
End Sub

Private Sub ContMore_Click()

    If IsNumeric(Me.TextBoxContracted) Then
        tmp = CLng(Me.TextBoxContracted)
        tmp = tmp + 1
        Me.TextBoxContracted = CStr(tmp)
    End If
End Sub

Private Sub ContMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxContracted) Then
        tmp = CLng(Me.TextBoxContracted)
        tmp = tmp + 10
        Me.TextBoxContracted = CStr(tmp)
    End If
End Sub

Private Sub OpenBPLess_Click()
    If IsNumeric(Me.TextBoxOpenBP) Then
        If CLng(Me.TextBoxOpenBP) > 0 Then
            tmp = CLng(Me.TextBoxOpenBP)
            tmp = tmp - 1
            Me.TextBoxOpenBP = CStr(tmp)
        End If
    End If
End Sub

Private Sub OpenBPLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxOpenBP) Then
        If CLng(Me.TextBoxOpenBP) > 9 Then
            tmp = CLng(Me.TextBoxOpenBP)
            tmp = tmp - 10
            Me.TextBoxOpenBP = CStr(tmp)
        End If
    End If
End Sub

Private Sub OpenBPMore_Click()
    If IsNumeric(Me.TextBoxOpenBP) Then
        tmp = CLng(Me.TextBoxOpenBP)
        tmp = tmp + 1
        Me.TextBoxOpenBP = CStr(tmp)
    End If
End Sub

Private Sub OpenBPMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxOpenBP) Then
        tmp = CLng(Me.TextBoxOpenBP)
        tmp = tmp + 10
        Me.TextBoxOpenBP = CStr(tmp)
    End If
End Sub

Private Sub PnocLess_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        If CLng(Me.TextBoxPNOC) > 0 Then
            tmp = CLng(Me.TextBoxPNOC)
            tmp = tmp - 1
            Me.TextBoxPNOC = CStr(tmp)
        End If
    End If
End Sub

Private Sub PnocLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxPNOC) Then
        If CLng(Me.TextBoxPNOC) > 9 Then
            tmp = CLng(Me.TextBoxPNOC)
            tmp = tmp - 10
            Me.TextBoxPNOC = CStr(tmp)
        End If
    End If
End Sub

Private Sub PnocMore_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        tmp = CLng(Me.TextBoxPNOC)
        tmp = tmp + 1
        Me.TextBoxPNOC = CStr(tmp)
    End If

End Sub
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------
Private Sub PnocMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxPNOC) Then
        tmp = CLng(Me.TextBoxPNOC)
        tmp = tmp + 10
        Me.TextBoxPNOC = CStr(tmp)
    End If
End Sub
