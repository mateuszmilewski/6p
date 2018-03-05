VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOseaScope 
   Caption         =   "Osea Scope"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   OleObjectBlob   =   "FormOseaScope.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOseaScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnImport_Click()
    ' MsgBox "not implemented yet!"
    Hide
    
    FormCatchWizard.ListBox1.Clear
    FormCatchWizard.ListBox1.MultiSelect = fmMultiSelectSingle
    
    Dim w As Workbook
    For Each w In Workbooks
        With FormCatchWizard.ListBox1
            .AddItem w.name
        End With
    Next w
    FormCatchWizard.czy_start_pochodzi_z_open_issues = False
    FormCatchWizard.BtnImportOpenIssues.Enabled = False
    FormCatchWizard.BtnJustImport.Enabled = False
    FormCatchWizard.BtnSubmit.Enabled = False
    FormCatchWizard.BtnOsea.Enabled = True
    FormCatchWizard.Show vbModeless

End Sub

Private Sub BtnSubmit_Click()


    SIXP.GlobalFooModule.gotoThisWorkbookMainA1

    ' text na guziki
    ' Global Const G_BTN_TEXT_ADD = "Dodaj"
    ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
    ' Hide
    inner_calc
    
    'run_FormMain Me.LabelTitle
End Sub

Private Sub change_col_I_in_MAIN_worksheet(ByRef r As Range)
    
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
                    rr.Offset(0, SIXP.e_main_last_update_on_osea - 1) = Trim(CStr(rr.Offset(0, 3)))
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
        Set r = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = Trim(arr(x))
        Next x
        
        recalc_total_textbox
        give_data_to_ranges r
        change_col_I_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm).Cells(1, 1)
        
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_I_in_MAIN_worksheet r
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
    r.Parent.Cells(r.Row, SIXP.e_osea_order_after_mrd) = CStr(Me.TextBoxAfterMRD)
    r.Parent.Cells(r.Row, SIXP.e_osea_order_confirmed) = CStr(Me.TextBoxConfirmed)
    r.Parent.Cells(r.Row, SIXP.e_osea_order_for_mrd) = CStr(Me.TextBoxForMRD)
    r.Parent.Cells(r.Row, SIXP.e_osea_order_on_stock) = CStr(Me.TextBoxOnStock)
    r.Parent.Cells(r.Row, SIXP.e_osea_order_open) = CStr(Me.TextBoxOPEN)
    r.Parent.Cells(r.Row, SIXP.e_osea_order_ordered) = CStr(Me.TextBoxOrdered)
    r.Parent.Cells(r.Row, SIXP.e_osea_order_total) = CStr(Me.TextBoxTotal)
End Sub

Private Sub recalc_total_textbox()

    Dim tmp_str_value As String
    Dim v As Long
    
    v = 0
    v = v + _
        CLng(SIXP.GlobalFooModule.global_cpz(CStr(Me.TextBoxAfterMRD))) + _
        CLng(SIXP.GlobalFooModule.global_cpz(CStr(Me.TextBoxConfirmed))) + _
        CLng(SIXP.GlobalFooModule.global_cpz(CStr(Me.TextBoxForMRD))) + _
        CLng(SIXP.GlobalFooModule.global_cpz(CStr(Me.TextBoxOnStock))) + _
        CLng(SIXP.GlobalFooModule.global_cpz(CStr(Me.TextBoxOPEN))) + _
        CLng(SIXP.GlobalFooModule.global_cpz(CStr(Me.TextBoxOrdered)))
        
        
    temp_str_textbox = CStr(v)
    
    Me.TextBoxTotal = _
        CStr(temp_str_textbox)
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




Private Sub AfterMRDLess_Click()

    If IsNumeric(Me.TextBoxAfterMRD) Then
        If CLng(Me.TextBoxAfterMRD) > 0 Then
            
            tmp = CLng(Me.TextBoxAfterMRD)
            tmp = tmp - 1
            Me.TextBoxAfterMRD = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub AfterMRDMore_Click()
    
    If IsNumeric(Me.TextBoxAfterMRD) Then
        tmp = CLng(Me.TextBoxAfterMRD)
        tmp = tmp + 1
        Me.TextBoxAfterMRD = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub


Private Sub AfterMRDLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If IsNumeric(Me.TextBoxAfterMRD) Then
        If CLng(Me.TextBoxAfterMRD) > 9 Then
            
            tmp = CLng(Me.TextBoxAfterMRD)
            tmp = tmp - 10
            Me.TextBoxAfterMRD = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 10
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub AfterMRDMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxAfterMRD) Then
        tmp = CLng(Me.TextBoxAfterMRD)
        tmp = tmp + 10
        Me.TextBoxAfterMRD = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 10
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub


Private Sub ConfirmedLess_Click()
    If IsNumeric(Me.TextBoxConfirmed) Then
        If CLng(Me.TextBoxConfirmed) > 0 Then
            
            tmp = CLng(Me.TextBoxConfirmed)
            tmp = tmp - 1
            Me.TextBoxConfirmed = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub



Private Sub ConfirmedMore_Click()
    If IsNumeric(Me.TextBoxConfirmed) Then
        tmp = CLng(Me.TextBoxConfirmed)
        tmp = tmp + 1
        Me.TextBoxConfirmed = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub


Private Sub ForMRDLess_Click()
    If IsNumeric(Me.TextBoxForMRD) Then
        If CLng(Me.TextBoxForMRD) > 0 Then
            
            tmp = CLng(Me.TextBoxForMRD)
            tmp = tmp - 1
            Me.TextBoxForMRD = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub ForMRDMore_Click()
    If IsNumeric(Me.TextBoxForMRD) Then
        tmp = CLng(Me.TextBoxForMRD)
        tmp = tmp + 1
        Me.TextBoxForMRD = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub


Private Sub OnStockLess_Click()
    If IsNumeric(Me.TextBoxOnStock) Then
        If CLng(Me.TextBoxOnStock) > 0 Then
            
            tmp = CLng(Me.TextBoxOnStock)
            tmp = tmp - 1
            Me.TextBoxOnStock = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub OnStockMore_Click()
    If IsNumeric(Me.TextBoxOnStock) Then
        tmp = CLng(Me.TextBoxOnStock)
        tmp = tmp + 1
        Me.TextBoxOnStock = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub



Private Sub OpenLess_Click()
    If IsNumeric(Me.TextBoxOPEN) Then
        If CLng(Me.TextBoxOPEN) > 0 Then
            
            tmp = CLng(Me.TextBoxOPEN)
            tmp = tmp - 1
            Me.TextBoxOPEN = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub OpenMore_Click()
    If IsNumeric(Me.TextBoxOPEN) Then
        tmp = CLng(Me.TextBoxOPEN)
        tmp = tmp + 1
        Me.TextBoxOPEN = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub

Private Sub OrderedLess_Click()
    If IsNumeric(Me.TextBoxOrdered) Then
        If CLng(Me.TextBoxOrdered) > 0 Then
            
            tmp = CLng(Me.TextBoxOrdered)
            tmp = tmp - 1
            Me.TextBoxOrdered = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub OrderedLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxOrdered) Then
        If CLng(Me.TextBoxOrdered) > 9 Then
            
            tmp = CLng(Me.TextBoxOrdered)
            tmp = tmp - 10
            Me.TextBoxOrdered = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 10
            'Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub OrderedMore_Click()
    If IsNumeric(Me.TextBoxOrdered) Then
        tmp = CLng(Me.TextBoxOrdered)
        tmp = tmp + 1
        Me.TextBoxOrdered = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub

Private Sub OrderedMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxOrdered) Then
        tmp = CLng(Me.TextBoxOrdered)
        tmp = tmp + 10
        Me.TextBoxOrdered = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 10
        'Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub

Private Sub TextBoxAfterMRD_Change()
    recalc_total_textbox
End Sub

Private Sub TextBoxConfirmed_Change()
    recalc_total_textbox
End Sub

Private Sub TextBoxForMRD_Change()
    recalc_total_textbox
End Sub

Private Sub TextBoxOnStock_Change()
    recalc_total_textbox
End Sub

Private Sub TextBoxOpen_Change()
    recalc_total_textbox
End Sub

Private Sub TextBoxOrdered_Change()
    recalc_total_textbox
End Sub

Private Sub TextBoxTotal_Change()
    recalc_total_textbox
End Sub

Private Sub TotLess_Click()
    If IsNumeric(Me.TextBoxTotal) Then
        If CLng(Me.TextBoxTotal) > 0 Then
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub TotLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxTotal) Then
        If CLng(Me.TextBoxTotal) > 9 Then
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 10
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
    recalc_total_textbox
End Sub

Private Sub TotMore_Click()
    If IsNumeric(Me.TextBoxTotal) Then
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub


Private Sub TotMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxTotal) Then
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 10
        Me.TextBoxTotal = CStr(tmp)
    End If
    recalc_total_textbox
End Sub

Private Sub UserForm_Activate()
    recalc_total_textbox
End Sub

Private Sub UserForm_Click()
    recalc_total_textbox
End Sub


Private Sub UserForm_Initialize()
    recalc_total_textbox
End Sub


Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    recalc_total_textbox
End Sub
