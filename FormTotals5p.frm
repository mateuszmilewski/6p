VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTotals5p 
   Caption         =   "TOTALS 5P"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   OleObjectBlob   =   "FormTotals5p.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTotals5p"
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

Private Sub change_col_J_in_MAIN_worksheet(ByRef r As Range)
    
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
                    rr.Offset(0, SIXP.e_main_last_update_on_totals - 1) = Trim(CStr(rr.Offset(0, 3)))
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
        Set r = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = arr(x)
        Next x
        
        
        give_data_to_ranges r
        change_col_J_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_J_in_MAIN_worksheet r
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
    r.Parent.Cells(r.Row, SIXP.e_5p_arrived) = CStr(Me.TextBoxArrived)
    r.Parent.Cells(r.Row, SIXP.e_5p_fma_eur) = CStr(Me.TextBoxFMAEUR)
    r.Parent.Cells(r.Row, SIXP.e_5p_fma_osea) = CStr(Me.TextBoxFmaOsea)
    r.Parent.Cells(r.Row, SIXP.e_5p_in_transit) = CStr(Me.TextBoxInTransit)
    r.Parent.Cells(r.Row, SIXP.e_5p_itdc) = CStr(Me.TextBoxITDC)
    r.Parent.Cells(r.Row, SIXP.e_5p_na) = CStr(Me.TextBoxNA)
    r.Parent.Cells(r.Row, SIXP.e_5p_no_ppap_status) = CStr(Me.TextBoxNoPPAP)
    r.Parent.Cells(r.Row, SIXP.e_5p_ordered) = CStr(Me.TextBoxOrdered)
    r.Parent.Cells(r.Row, SIXP.e_5p_pnoc) = CStr(Me.TextBoxPNOC)
    r.Parent.Cells(r.Row, SIXP.e_5p_ppap_status) = CStr(Me.TextBoxPPAP)
    r.Parent.Cells(r.Row, SIXP.e_5p_total) = CStr(Me.TextBoxTotal)
End Sub

' textboxes with qtyies bedzie w sumie 22 procedury wiec bierz sie do roboty
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


Private Sub ArrivedLess_Click()
    If IsNumeric(Me.TextBoxArrived) Then
        If CLng(Me.TextBoxArrived) > 0 Then
            tmp = CLng(Me.TextBoxArrived)
            tmp = tmp - 1
            Me.TextBoxArrived = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub ArrivedMore_Click()
    If IsNumeric(Me.TextBoxArrived) Then
        tmp = CLng(Me.TextBoxArrived)
        tmp = tmp + 1
        Me.TextBoxArrived = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub FMAEURLEss_Click()
    If IsNumeric(Me.TextBoxFMAEUR) Then
        If CLng(Me.TextBoxFMAEUR) > 0 Then
            tmp = CLng(Me.TextBoxFMAEUR)
            tmp = tmp - 1
            Me.TextBoxFMAEUR = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub FMAEURMore_Click()
    If IsNumeric(Me.TextBoxFMAEUR) Then
        tmp = CLng(Me.TextBoxFMAEUR)
        tmp = tmp + 1
        Me.TextBoxFMAEUR = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub FmaOseaLess_Click()
    If IsNumeric(Me.TextBoxFmaOsea) Then
        If CLng(Me.TextBoxFmaOsea) > 0 Then
            tmp = CLng(Me.TextBoxFmaOsea)
            tmp = tmp - 1
            Me.TextBoxFmaOsea = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub FmaOseaMore_Click()
    If IsNumeric(Me.TextBoxFmaOsea) Then
        tmp = CLng(Me.TextBoxFmaOsea)
        tmp = tmp + 1
        Me.TextBoxFmaOsea = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub InTransitLess_Click()
    If IsNumeric(Me.TextBoxInTransit) Then
        If CLng(Me.TextBoxInTransit) > 0 Then
            tmp = CLng(Me.TextBoxInTransit)
            tmp = tmp - 1
            Me.TextBoxInTransit = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub InTransitMore_Click()
    If IsNumeric(Me.TextBoxInTransit) Then
        tmp = CLng(Me.TextBoxInTransit)
        tmp = tmp + 1
        Me.TextBoxInTransit = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub ITDCLess_Click()
    If IsNumeric(Me.TextBoxITDC) Then
        If CLng(Me.TextBoxITDC) > 0 Then
            tmp = CLng(Me.TextBoxITDC)
            tmp = tmp - 1
            Me.TextBoxITDC = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub ITDCMore_Click()
    If IsNumeric(Me.TextBoxITDC) Then
        tmp = CLng(Me.TextBoxITDC)
        tmp = tmp + 1
        Me.TextBoxITDC = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub NALess_Click()
    If IsNumeric(Me.TextBoxNA) Then
        If CLng(Me.TextBoxNA) > 0 Then
            tmp = CLng(Me.TextBoxNA)
            tmp = tmp - 1
            Me.TextBoxNA = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub NAMore_Click()
    If IsNumeric(Me.TextBoxNA) Then
        tmp = CLng(Me.TextBoxNA)
        tmp = tmp + 1
        Me.TextBoxNA = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub NoPPAPLess_Click()
    If IsNumeric(Me.TextBoxNoPPAP) Then
        If CLng(Me.TextBoxNoPPAP) > 0 Then
            tmp = CLng(Me.TextBoxNoPPAP)
            tmp = tmp - 1
            Me.TextBoxNoPPAP = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub NoPPAPMore_Click()
    If IsNumeric(Me.TextBoxNoPPAP) Then
        tmp = CLng(Me.TextBoxNoPPAP)
        tmp = tmp + 1
        Me.TextBoxNoPPAP = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub OrderedLess_Click()
    If IsNumeric(Me.TextBoxOrdered) Then
        If CLng(Me.TextBoxOrdered) > 0 Then
            tmp = CLng(Me.TextBoxOrdered)
            tmp = tmp - 1
            Me.TextBoxOrdered = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub OrderedMore_Click()
    If IsNumeric(Me.TextBoxOrdered) Then
        tmp = CLng(Me.TextBoxOrdered)
        tmp = tmp + 1
        Me.TextBoxOrdered = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub PnocLess_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        If CLng(Me.TextBoxPNOC) > 0 Then
            tmp = CLng(Me.TextBoxPNOC)
            tmp = tmp - 1
            Me.TextBoxPNOC = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub PnocMore_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        tmp = CLng(Me.TextBoxPNOC)
        tmp = tmp + 1
        Me.TextBoxPNOC = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub PPAPLess_Click()
    If IsNumeric(Me.TextBoxPPAP) Then
        If CLng(Me.TextBoxPPAP) > 0 Then
            tmp = CLng(Me.TextBoxPPAP)
            tmp = tmp - 1
            Me.TextBoxPPAP = CStr(tmp)
            
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub PPAPMore_Click()
    If IsNumeric(Me.TextBoxPPAP) Then
        tmp = CLng(Me.TextBoxPPAP)
        tmp = tmp + 1
        Me.TextBoxPPAP = CStr(tmp)
        
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub

Private Sub TotalLess_Click()
    If IsNumeric(Me.TextBoxTotal) Then
        If CLng(Me.TextBoxTotal) > 0 Then
            tmp = CLng(Me.TextBoxTotal)
            tmp = tmp - 1
            Me.TextBoxTotal = CStr(tmp)
        End If
    End If
End Sub

Private Sub TotalMore_Click()
    If IsNumeric(Me.TextBoxTotal) Then
        tmp = CLng(Me.TextBoxTotal)
        tmp = tmp + 1
        Me.TextBoxTotal = CStr(tmp)
    End If
End Sub
