VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormTotals5p 
   Caption         =   "TOTALS 5P"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   OleObjectBlob   =   "FormTotals5p.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormTotals5p"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
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


Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnSubmit_Click()

    ' text na guziki
    ' Global Const G_BTN_TEXT_ADD = "Dodaj"
    ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
    'Hide
    inner_calc
    
    ' run_FormMain Me.LabelTitle
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


    Dim r As Range
    
    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For X = 0 To 3
            r.Offset(0, X) = Trim(arr(X))
        Next X
        
        
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
    r.Parent.Cells(r.Row, SIXP.e_5p_future) = CStr(Me.TextBoxFuture)
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
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub ArrivedMore_Click()
    If IsNumeric(Me.TextBoxArrived) Then
        tmp = CLng(Me.TextBoxArrived)
        tmp = tmp + 1
        Me.TextBoxArrived = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub BtnZeruj_Click()
    wyzeruj_all
End Sub

Private Sub FMAEURLEss_Click()
    If IsNumeric(Me.TextBoxFMAEUR) Then
        If CLng(Me.TextBoxFMAEUR) > 0 Then
            tmp = CLng(Me.TextBoxFMAEUR)
            tmp = tmp - 1
            Me.TextBoxFMAEUR = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub FMAEURMore_Click()
    If IsNumeric(Me.TextBoxFMAEUR) Then
        tmp = CLng(Me.TextBoxFMAEUR)
        tmp = tmp + 1
        Me.TextBoxFMAEUR = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub FmaOseaLess_Click()
    If IsNumeric(Me.TextBoxFmaOsea) Then
        If CLng(Me.TextBoxFmaOsea) > 0 Then
            tmp = CLng(Me.TextBoxFmaOsea)
            tmp = tmp - 1
            Me.TextBoxFmaOsea = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub FmaOseaMore_Click()
    If IsNumeric(Me.TextBoxFmaOsea) Then
        tmp = CLng(Me.TextBoxFmaOsea)
        tmp = tmp + 1
        Me.TextBoxFmaOsea = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub FutureLess_Click()
    If IsNumeric(Me.TextBoxFuture) Then
        If CLng(Me.TextBoxFuture) > 0 Then
            tmp = CLng(Me.TextBoxFuture)
            tmp = tmp - 1
            Me.TextBoxFuture = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub FutureMore_Click()
    If IsNumeric(Me.TextBoxFuture) Then
        tmp = CLng(Me.TextBoxFuture)
        tmp = tmp + 1
        Me.TextBoxFuture = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub InTransitLess_Click()
    If IsNumeric(Me.TextBoxInTransit) Then
        If CLng(Me.TextBoxInTransit) > 0 Then
            tmp = CLng(Me.TextBoxInTransit)
            tmp = tmp - 1
            Me.TextBoxInTransit = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub InTransitMore_Click()
    If IsNumeric(Me.TextBoxInTransit) Then
        tmp = CLng(Me.TextBoxInTransit)
        tmp = tmp + 1
        Me.TextBoxInTransit = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub ITDCLess_Click()
    If IsNumeric(Me.TextBoxITDC) Then
        If CLng(Me.TextBoxITDC) > 0 Then
            tmp = CLng(Me.TextBoxITDC)
            tmp = tmp - 1
            Me.TextBoxITDC = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub ITDCMore_Click()
    If IsNumeric(Me.TextBoxITDC) Then
        tmp = CLng(Me.TextBoxITDC)
        tmp = tmp + 1
        Me.TextBoxITDC = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub



Private Sub Label16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.TextBoxH1.Value = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("h1")
    
    ' jeszcze szybkie sprawdzenie sum:
    ' --------------------------------------------------------------------
    
    If IsNumeric(Me.TextBoxH1) And IsNumeric(Me.TextBoxTotal) Then
    
        If CLng(Me.TextBoxH1.Value) = CLng(Me.TextBoxTotal.Value) Then
            Me.TextBoxH1.BackColor = RGB(255, 255, 255)
        Else
        
            If CLng(Me.TextBoxH1.Value) < CLng(Me.TextBoxTotal.Value) Then
                Me.TextBoxH1.BackColor = RGB(255, 0, 0)
            ElseIf CLng(Me.TextBoxH1.Value) > CLng(Me.TextBoxTotal.Value) Then
                Me.TextBoxH1.BackColor = RGB(255, 255, 0)
            Else
                Me.TextBoxH1.BackColor = RGB(255, 0, 0)
            End If
        End If
    Else
        Me.TextBoxH1.BackColor = RGB(255, 0, 0)
    End If
    
    ' --------------------------------------------------------------------
End Sub

Private Sub NALess_Click()
    If IsNumeric(Me.TextBoxNA) Then
        If CLng(Me.TextBoxNA) > 0 Then
            tmp = CLng(Me.TextBoxNA)
            tmp = tmp - 1
            Me.TextBoxNA = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub NAMore_Click()
    If IsNumeric(Me.TextBoxNA) Then
        tmp = CLng(Me.TextBoxNA)
        tmp = tmp + 1
        Me.TextBoxNA = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub NoPPAPLess_Click()
    If IsNumeric(Me.TextBoxNoPPAP) Then
        If CLng(Me.TextBoxNoPPAP) > 0 Then
            tmp = CLng(Me.TextBoxNoPPAP)
            tmp = tmp - 1
            Me.TextBoxNoPPAP = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub NoPPAPMore_Click()
    If IsNumeric(Me.TextBoxNoPPAP) Then
        tmp = CLng(Me.TextBoxNoPPAP)
        tmp = tmp + 1
        Me.TextBoxNoPPAP = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
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
            to_je_synchro
        End If
    End If
End Sub

Private Sub OrderedMore_Click()
    If IsNumeric(Me.TextBoxOrdered) Then
        tmp = CLng(Me.TextBoxOrdered)
        tmp = tmp + 1
        Me.TextBoxOrdered = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub PnocLess_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        If CLng(Me.TextBoxPNOC) > 0 Then
            tmp = CLng(Me.TextBoxPNOC)
            tmp = tmp - 1
            Me.TextBoxPNOC = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub PnocMore_Click()
    If IsNumeric(Me.TextBoxPNOC) Then
        tmp = CLng(Me.TextBoxPNOC)
        tmp = tmp + 1
        Me.TextBoxPNOC = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

Private Sub PPAPLess_Click()
    If IsNumeric(Me.TextBoxPPAP) Then
        If CLng(Me.TextBoxPPAP) > 0 Then
            tmp = CLng(Me.TextBoxPPAP)
            tmp = tmp - 1
            Me.TextBoxPPAP = CStr(tmp)
            
            'tmp = CLng(Me.TextBoxTotal)
            'tmp = tmp - 1
            'Me.TextBoxTotal = CStr(tmp)
            to_je_synchro
        End If
    End If
End Sub

Private Sub PPAPMore_Click()
    If IsNumeric(Me.TextBoxPPAP) Then
        tmp = CLng(Me.TextBoxPPAP)
        tmp = tmp + 1
        Me.TextBoxPPAP = CStr(tmp)
        
        'tmp = CLng(Me.TextBoxTotal)
        'tmp = tmp + 1
        'Me.TextBoxTotal = CStr(tmp)
        to_je_synchro
    End If
End Sub

'Private Sub TotalLess_Click()
'    If IsNumeric(Me.TextBoxTotal) Then
'        If CLng(Me.TextBoxTotal) > 0 Then
'            tmp = CLng(Me.TextBoxTotal)
'            tmp = tmp - 1
'            Me.TextBoxTotal = CStr(tmp)
'        End If
'    End If
'End Sub
'
'Private Sub TotalMore_Click()
'    If IsNumeric(Me.TextBoxTotal) Then
'        tmp = CLng(Me.TextBoxTotal)
'        tmp = tmp + 1
'        Me.TextBoxTotal = CStr(tmp)
'    End If
'End Sub

Private Sub wyzeruj_all()
    
    Dim c As Control
    Dim tb As Control
    
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            Set tb = c
            tb.Value = 0
        End If
    Next c
End Sub


Private Sub to_je_synchro()
    
    Dim c As Control
    Dim tb As Control
    
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            Set tb = c
            
            If tb.Enabled = True Then
                If tb.Value = "" Then
                    
                    tb.Value = 0
                End If
            End If
        End If
    Next c
    
    Me.TextBoxTotal = 0
    
    Dim suma As Long
    suma = 0
    
    pref = "TextBox"
    
    For Each c In Me.Controls
        If TypeName(c) = CStr(pref) Then
            Set tb = c
            
            If tb.Enabled = True Then
            
                If tb.Name = pref & "NA" Or _
                    tb.Name = pref & "ITDC" Or _
                    tb.Name = pref & "PNOC" Or _
                    tb.Name = pref & "FMAEUR" Or _
                    tb.Name = pref & "FmaOsea" Then
                
                'If tb.Name = pref & "NA" Or _
                '    tb.Name = pref & "FMAEUR" Or _
                '    tb.Name = pref & "FmaOsea" Then
                    
                        
                        suma = suma + CLng(tb.Value)
                End If
            End If
        End If
    Next c
    
    Me.TextBoxTotal = suma
    
    If Me.TextBoxTotal > 0 Then
    ' dodatkowo pokoloruj gdy wartosci sa podejrzanie wyzsze niz total
    ' --------------------------------------------------------------------
    
    ' to jest podjerzanie za duzo
    If CLng(Me.TextBoxPPAP.Value) + CLng(Me.TextBoxNoPPAP.Value) > CLng(Me.TextBoxTotal) Then
        
        Me.TextBoxPPAP.BackColor = RGB(255, 0, 0)
        Me.TextBoxNoPPAP.BackColor = RGB(255, 0, 0)
    Else
    
        If CLng(Me.TextBoxPPAP.Value) + CLng(Me.TextBoxNoPPAP.Value) < CLng(Me.TextBoxTotal) Then
        
            Me.TextBoxPPAP.BackColor = RGB(255, 255, 0)
            Me.TextBoxNoPPAP.BackColor = RGB(255, 255, 0)
        Else
            Me.TextBoxPPAP.BackColor = RGB(255, 255, 255)
            Me.TextBoxNoPPAP.BackColor = RGB(255, 255, 255)
        End If
    End If
    
    
    ' tranzyty
    ' --------------------------------------------------------------------
    suma_tranzytow = CLng(Me.TextBoxArrived.Value) + CLng(Me.TextBoxInTransit.Value) + CLng(Me.TextBoxFuture.Value)
    If CLng(suma_tranzytow) > CLng(Me.TextBoxTotal) Then
    
        Me.TextBoxArrived.BackColor = RGB(255, 0, 0)
        Me.TextBoxInTransit.BackColor = RGB(255, 0, 0)
        Me.TextBoxFuture.BackColor = RGB(255, 0, 0)
    Else
    
        If CLng(suma_tranzytow) < CLng(Me.TextBoxTotal) Then
            Me.TextBoxArrived.BackColor = RGB(255, 255, 0)
            Me.TextBoxInTransit.BackColor = RGB(255, 255, 0)
            Me.TextBoxFuture.BackColor = RGB(255, 255, 0)
        Else
            Me.TextBoxArrived.BackColor = RGB(255, 255, 255)
            Me.TextBoxInTransit.BackColor = RGB(255, 255, 255)
            Me.TextBoxFuture.BackColor = RGB(255, 255, 255)
        End If
    End If
    
    
    ' ordered
    ' --------------------------------------------------------------------
    If CLng(Me.TextBoxOrdered.Value) > CLng(Me.TextBoxTotal) Then
        Me.TextBoxOrdered.BackColor = RGB(255, 0, 0)
    Else
        Me.TextBoxOrdered.BackColor = RGB(255, 255, 255)
    End If
    ' --------------------------------------------------------------------
    
    
    
    ' jeszcze szybkie sprawdzenie sum:
    ' --------------------------------------------------------------------
    
    If IsNumeric(Me.TextBoxH1) And IsNumeric(Me.TextBoxTotal) Then
    
        If CLng(Me.TextBoxH1.Value) = CLng(Me.TextBoxTotal.Value) Then
            Me.TextBoxH1.BackColor = RGB(255, 255, 255)
        Else
            If CLng(Me.TextBoxH1.Value) < CLng(Me.TextBoxTotal.Value) Then
                Me.TextBoxH1.BackColor = RGB(255, 0, 0)
            ElseIf CLng(Me.TextBoxH1.Value) > CLng(Me.TextBoxTotal.Value) Then
                Me.TextBoxH1.BackColor = RGB(255, 255, 0)
            Else
                Me.TextBoxH1.BackColor = RGB(255, 0, 0)
            End If
        End If
    Else
        Me.TextBoxH1.BackColor = RGB(255, 0, 0)
    End If
    
    ' --------------------------------------------------------------------
    End If
    
End Sub


Private Sub TextBoxArrived_Change()
    to_je_synchro
End Sub

Private Sub TextBoxFMAEUR_Change()
    to_je_synchro
End Sub

Private Sub TextBoxFmaOsea_Change()
    to_je_synchro
End Sub

Private Sub TextBoxFuture_Change()
    to_je_synchro
End Sub




Private Sub TextBoxInTransit_Change()
    to_je_synchro
End Sub

Private Sub TextBoxITDC_Change()
    to_je_synchro
End Sub

Private Sub TextBoxNA_Change()
    to_je_synchro
End Sub

Private Sub TextBoxNoPPAP_Change()
    to_je_synchro
End Sub

Private Sub TextBoxOrdered_Change()
    to_je_synchro
End Sub

Private Sub TextBoxPNOC_Change()
    to_je_synchro
End Sub

Private Sub TextBoxPPAP_Change()
    to_je_synchro
End Sub

Private Sub TextBoxTotal_Change()
    ' to_je_synchro
End Sub

Private Sub TryWizardBtn_Click()
    ' sub odpowiadajacy za sciaganie danych z wizard buff worksheet
    
    
    ' MsgBox "not implemented yet!"
    
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    '3: MRD
    '4: BUILD START
    '5: BUILD END
    '6: BOM
    '7: PPAP GATE
    
    total_num = 0
    If IsNumeric(buff.Cells(3, 1)) Then
        total_num = CLng(buff.Cells(3, 1))
    End If
    
    With buff
        
        
        ' porcja zwiazana z totalami eur i osea
        ' ---------------------------------------------------------------
        osea_v = 0
        eur_v = 0
        
        If IsNumeric(.Cells(32, 1)) Then osea_v = .Cells(32, 1)
        If IsNumeric(.Cells(32, 2)) Then eur_v = .Cells(32, 2)
        ' ---------------------------------------------------------------
        
        
        ' porcja total z perspektywy transportow
        ' ---------------------------------------------------------------
        arrived = 0
        in_t = 0
        future = 0
        
        If IsNumeric(.Cells(37, 1)) Then arrived = .Cells(37, 1)
        If IsNumeric(.Cells(37, 2)) Then in_t = .Cells(37, 2)
        If IsNumeric(.Cells(37, 3)) Then future = .Cells(37, 3)
        ' ---------------------------------------------------------------
        
        ' porcja PNOC
        ' nieco bardziej zlozona bo trzeba jeszcze znalezc slowo
        ' klucz w wierszy drugim, liczba w wierszu 3
        ' ---------------------------------------------------------------
        pnoc_v = 0
        pnoc_v = ktora_kolumna_to__foo(.Cells(2, 1), "PNOC")
        ' ---------------------------------------------------------------
        
        
        ' ---------------------------------------------------------------
        itdc_v = 0
        itdc_v = ktora_kolumna_to__foo(.Cells(2, 1), "ITDC")
        ' ---------------------------------------------------------------
        
        ' ---------------------------------------------------------------
        mpc_v = 0
        mpc_v = ktora_kolumna_to__foo(.Cells(2, 1), "MPC")
        ' ---------------------------------------------------------------
        
        ' ---------------------------------------------------------------
        ordered_v = 0
        ordered_v = ktora_kolumna_to__foo(.Cells(41, 1), "OK")
        ' ---------------------------------------------------------------
        
        
        ' PPAP
        ' na ppap skladaja sie ok i nok i koncepcja obliczeniowa jest nieco bardziej skomplikowana
        ' bedzie trzeba sie posilkowac lista w rejestrze, ktore stringi sa ok, a ktore sa nok
        ' ---------------------------------------------------------------
        
        ppap_ok = 0
        ppap_nok = 0
        
        ppap_ok = get_ppaps_foo(.Cells(6, 1), E_PPAP_OK)
        ppap_nok = get_ppaps_foo(.Cells(6, 1), E_PPAP_NOK)
        
        ' ---------------------------------------------------------------
    End With
    
    
    ' teraz wyliczymy NA na podstawie roznicy totala z juz zaciagnietych danych
    na_v = 0
    '!
    
    na_v = CLng(total_num - pnoc_v - itdc_v - osea_v - eur_v)
    
    If CLng(na_v) < 0 Then
        na_v = 0
    End If
    
    
    Me.TextBoxArrived = arrived
    Me.TextBoxFMAEUR = eur_v
    Me.TextBoxFmaOsea = osea_v
    Me.TextBoxFuture = future
    Me.TextBoxInTransit = in_t
    Me.TextBoxITDC = itdc_v
    Me.TextBoxNA = na_v
    Me.TextBoxNoPPAP = ppap_nok
    Me.TextBoxOrdered = ordered_v
    Me.TextBoxPNOC = pnoc_v
    Me.TextBoxPPAP = ppap_ok
    Me.TextBoxTotal = total_num
    
    Me.TextBoxH1 = buff.Range("H1")
    
    ' jeszcze szybkie sprawdzenie sum:
    ' --------------------------------------------------------------------
    
    If IsNumeric(Me.TextBoxH1) And IsNumeric(Me.TextBoxTotal) Then
    
        If CLng(Me.TextBoxH1.Value) = CLng(Me.TextBoxTotal.Value) Then
            Me.TextBoxH1.BackColor = RGB(255, 255, 255)
        Else
            Me.TextBoxH1.BackColor = RGB(255, 0, 0)
        End If
    Else
        Me.TextBoxH1.BackColor = RGB(255, 0, 0)
    End If
    
    ' --------------------------------------------------------------------
    
End Sub

Private Function get_ppaps_foo(r As Range, e As E_PPAP) As Long
    get_ppaps_foo = 0
    
    
    ' 2 petle
    ' jedna po wiz buff
    ' druga po register list
    
    ' range wizz buff
    Dim rwb As Range
    ' range register
    Dim rr As Range, tmp_r As Range
    
    Set rwb = r
    
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("J2")
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range(rr, rr.End(xlDown))
    
    Do
        For Each tmp_r In rr
            If Trim(tmp_r) = Trim(rwb) Then
                If CLng(tmp_r.Offset(0, 1)) = CLng(e) Then
                    
                    ' ----------------------------------------------------------------
                    ' ----------------------------------------------------------------
                    ''
                    '
                    get_ppaps_foo = get_ppaps_foo + CLng(rwb.Offset(1, 0))
                    '
                    ''
                    ' ----------------------------------------------------------------
                    ' ----------------------------------------------------------------
                End If
            End If
        Next tmp_r
        Set rwb = rwb.Offset(0, 1)
    Loop Until Trim(rwb) = ""
    
End Function

Private Function ktora_kolumna_to__foo(r As Range, txt As String) As Long
    
    ktora_kolumna_to__foo = 0
    
    Do
        If Trim(r) = CStr(txt) Then
            ktora_kolumna_to__foo = CLng(r.Offset(1, 0))
            Exit Function
        End If
        
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
End Function

Private Sub UserForm_Activate()
    to_je_synchro
End Sub

Private Sub UserForm_Click()
    to_je_synchro
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    to_je_synchro
End Sub
