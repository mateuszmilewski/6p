VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormContractedPNOC 
   Caption         =   "FormContractedPNOC"
   ClientHeight    =   4800
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
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Private walidator As Validator


Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnImport_Click()
    
    
    ' importujemy dane z resp z buffa
    '-----------------------------------
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    
    ' rngv - range values, range labels
    Dim rngv As Range, rngl As Range, total_total
    Set rngv = buff.Cells(3, 1)
    Set rngl = buff.Cells(2, 1)
    Set total_total = buff.Range("B4")
    
    
    With buff

        Me.TextBoxPNOC.Value = CStr(SIXP.GetDataFromWizardBufferModule.get_all_values("PNOC", rngl))
        ' total total
        Me.TextBox1.Value = CStr(CLng(total_total.Value)) ' - CLng(Me.TextBoxPNOC.Value))
        Me.TextBoxContracted = CStr(CLng(Me.TextBox1.Value) - CLng(Me.TextBoxPNOC))
    End With
    
    
    '-----------------------------------
    jaka_kolwiek_zmiana_nastapila
End Sub




Private Sub jaka_kolwiek_zmiana_nastapila()


    With Me
    
        If Trim(.TextBoxActionableFMA.Value) = "" Then
            .TextBoxActionableFMA.Value = "0"
        End If
        
        
        If Trim(.TextBoxContracted.Value) = "" Then
            .TextBoxContracted.Value = "0"
        End If
        
        If Trim(.TextBoxOpenBP.Value) = "" Then
            .TextBoxOpenBP.Value = "0"
        End If
        
        If Trim(.TextBoxPNOC.Value) = "" Then
            .TextBoxPNOC.Value = "0"
        End If
    
    
        If IsNumeric(.TextBox1.Value) Then
            suma = CLng(.TextBoxActionableFMA) + CLng(.TextBoxContracted) + CLng(.TextBoxOpenBP) + CLng(.TextBoxPNOC)
            
            
            If CLng(suma) < CLng(.TextBox1.Value) Then
                .TextBox1.BackColor = RGB(255, 255, 0)
            ElseIf CLng(suma) = CLng(.TextBox1.Value) Then
                .TextBox1.BackColor = RGB(0, 255, 0)
            Else
                .TextBox1.BackColor = RGB(255, 0, 0)
            End If
        End If
    End With
End Sub

Private Sub BtnSubmit_Click()

    Set walidator = New Validator
    With walidator
        .dodajDoKolekcji Me.TextBox1, .pStr_checkIfNumber ' to jest suma - narazie bede to sprawdzal
        .dodajDoKolekcji Me.TextBoxActionableFMA, .pStr_checkIfNumber
        .dodajDoKolekcji Me.TextBoxContracted, .pStr_checkIfNumber
        .dodajDoKolekcji Me.TextBoxOpenBP, .pStr_checkIfNumber
        .dodajDoKolekcji Me.TextBoxPNOC, .pStr_checkIfNumber
        
        .run
    End With


    If walidator.pass Then


        SIXP.GlobalFooModule.gotoThisWorkbookMainA1
    
        ' text na guziki
        ' Global Const G_BTN_TEXT_ADD = "Dodaj"
        ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
        'Hide
        inner_calc
        
        ' run_FormMain Me.LabelTitle
        
        
        If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
            Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT
        End If
        
    Else
        MsgBox "Validation failed!"
    End If
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
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = Trim(arr(x))
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
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub ActionableFMALess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxActionableFMA) Then
        If CLng(Me.TextBoxActionableFMA) > 9 Then
            tmp = CLng(Me.TextBoxActionableFMA)
            tmp = tmp - 10
            Me.TextBoxActionableFMA = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub ActionableFMAMore_Click()
    If IsNumeric(Me.TextBoxActionableFMA) Then
        tmp = CLng(Me.TextBoxActionableFMA)
        tmp = tmp + 1
        Me.TextBoxActionableFMA = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub ActionableFMAMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxActionableFMA) Then
        tmp = CLng(Me.TextBoxActionableFMA)
        tmp = tmp + 10
        Me.TextBoxActionableFMA = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub ContLess_Click()
    If IsNumeric(Me.TextBoxContracted) Then
        If CLng(Me.TextBoxContracted) > 0 Then
            tmp = CLng(Me.TextBoxContracted)
            tmp = tmp - 1
            Me.TextBoxContracted = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub



Private Sub ContLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxContracted) Then
        If CLng(Me.TextBoxContracted) > 9 Then
            tmp = CLng(Me.TextBoxContracted)
            tmp = tmp - 10
            Me.TextBoxContracted = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub ContMore_Click()

    If IsNumeric(Me.TextBoxContracted) Then
        tmp = CLng(Me.TextBoxContracted)
        tmp = tmp + 1
        Me.TextBoxContracted = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub ContMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxContracted) Then
        tmp = CLng(Me.TextBoxContracted)
        tmp = tmp + 10
        Me.TextBoxContracted = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub OpenBPLess_Click()
    If IsNumeric(Me.TextBoxOpenBP) Then
        If CLng(Me.TextBoxOpenBP) > 0 Then
            tmp = CLng(Me.TextBoxOpenBP)
            tmp = tmp - 1
            Me.TextBoxOpenBP = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub OpenBPLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxOpenBP) Then
        If CLng(Me.TextBoxOpenBP) > 9 Then
            tmp = CLng(Me.TextBoxOpenBP)
            tmp = tmp - 10
            Me.TextBoxOpenBP = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub OpenBPMore_Click()
    If IsNumeric(Me.TextBoxOpenBP) Then
        tmp = CLng(Me.TextBoxOpenBP)
        tmp = tmp + 1
        Me.TextBoxOpenBP = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub OpenBPMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxOpenBP) Then
        tmp = CLng(Me.TextBoxOpenBP)
        tmp = tmp + 10
        Me.TextBoxOpenBP = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub PnocLess_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        If CLng(Me.TextBoxPNOC) > 0 Then
            tmp = CLng(Me.TextBoxPNOC)
            tmp = tmp - 1
            Me.TextBoxPNOC = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub PnocLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxPNOC) Then
        If CLng(Me.TextBoxPNOC) > 9 Then
            tmp = CLng(Me.TextBoxPNOC)
            tmp = tmp - 10
            Me.TextBoxPNOC = CStr(tmp)
        End If
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub PnocMore_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        tmp = CLng(Me.TextBoxPNOC)
        tmp = tmp + 1
        Me.TextBoxPNOC = CStr(tmp)
    End If
    jaka_kolwiek_zmiana_nastapila
End Sub
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------
Private Sub PnocMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxPNOC) Then
        tmp = CLng(Me.TextBoxPNOC)
        tmp = tmp + 10
        Me.TextBoxPNOC = CStr(tmp)
    End If
    
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub TextBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub TextBoxActionableFMA_Change()
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub TextBoxContracted_Change()
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub TextBoxOpenBP_Change()
    jaka_kolwiek_zmiana_nastapila
End Sub

Private Sub TextBoxPNOC_Change()
    jaka_kolwiek_zmiana_nastapila
End Sub
