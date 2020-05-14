VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOpenIssues 
   Caption         =   "Open Issues Form"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   OleObjectBlob   =   "FormOpenIssues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOpenIssues"
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


Private r As Range

Private Sub BtnDelete_Click()



    If sth_is_selected() Then
        
        ans = MsgBox("Chcesz usunac ten element?", vbYesNo)
        
        If ans = vbYes Then
    
            ' Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Cells(1, 1)
            
            v = Me.ListBox1.Value
            adr = adr_txt_parsed_from_selected_v_from_listbox(v)
            
            Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Range(CStr(adr))
            
            If CStr(Me.LabelTitle) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                    ' -----------------------------------------------------------------------------------------------------------
                    r.EntireRow.Delete xlUp
                    
                    inner_submit E_OI_DELETE
                    trzeba_jeszcze_usunac_wpis_w_arkuszu_main_jesli_nie_ma_wiecej_wpisow CStr(Me.LabelTitle)
                    
                    Me.Repaint
                    
                    Me.TextBoxComment = ""
                    Me.TextBoxDelivery = ""
                    Me.TextBoxNoOfPNs = ""
                    Me.TextBoxPartSupplier = ""
                    Me.ComboBoxStatus.Value = ""
                    Me.CheckBoxVisible.Value = True
                    ' -----------------------------------------------------------------------------------------------------------
            Else
                MsgBox "cos poszlo fest nie tak z pobraniem adresu komorki z open issues - program zatrzymal sie krytycznie"
                End
            End If
        Else
            MsgBox "zadna akcja nie zostala wykonana"
        End If
    Else
        MsgBox "nic nie wybrano!"
    End If
End Sub

Private Function sth_is_selected() As Boolean
    
    sth_is_selected = False
    
    For x = 0 To Me.ListBox1.ListCount - 1
        
        If Me.ListBox1.Selected(x) Then
            sth_is_selected = True
            Exit Function
        End If
    Next x
End Function

Private Sub trzeba_jeszcze_usunac_wpis_w_arkuszu_main_jesli_nie_ma_wiecej_wpisow(lbl_title As String)


    If Me.ListBox1.ListCount = 0 Then
    
    
        Dim mr As Range
        Set mr = ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Range("a1")
        
        Do
            If sprawdz_pierwsze_cztery_kolumny(mr, lbl_title) Then
            
                mr.Offset(0, SIXP.e_main_last_update_on_open_issues - 1).Value = ""
                
            End If
            Set mr = mr.Offset(1, 0)
        Loop Until mr.Row = 10000 ' heurystycznie dla przypadkow blankow po srodku tabeli
    End If
End Sub

Private Function sprawdz_pierwsze_cztery_kolumny(mr As Range, lbl_title As String) As Boolean
    sprawdz_pierwsze_cztery_kolumny = False
    
    If CStr(lbl_title) = _
        CStr(Trim(mr) & ", " & Trim(mr.Offset(0, 1)) & ", " & Trim(mr.Offset(0, 2)) & ", " & Trim(mr.Offset(0, 3))) Then
        
        
        sprawdz_pierwsze_cztery_kolumny = True
    End If
    
End Function

Private Sub BtnDodajNowy_Click()
    
    inner_submit E_OI_CREATE

End Sub

Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnImport_Click()
    
    ' MsgBox "not implemented yet!"
    ' Exit Sub
    
    ' funkcja importu - calkiem wazna
    ' ---------------------------------------------------------
    
    ' wczesniej ten msgbox mial byc jako tako masowy
    ' jednak z perspektywy designu calej apki nie moge tak zrobic
    ' zatem zatem: tutaj tylko dodaje dane do tego formularza
    ans = MsgBox("Czy chcesz zaciagnac jednorazowo informacje z otwartego pliku Wizard?", vbOKCancel, "Wizard Synchro")
    
    If ans = vbOK Then
        Hide
        
        
        ' usuniecie danych z wizard buff
        ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("a1:zz1000").Clear
        
        FormCatchWizard.ListBox1.Clear
        FormCatchWizard.ListBox1.MultiSelect = fmMultiSelectSingle
        
        For Each w In Workbooks
            With FormCatchWizard.ListBox1
                .addItem w.name
            End With
        Next w
        FormCatchWizard.czy_start_pochodzi_z_open_issues = True
        FormCatchWizard.wez_text_z_open_issues_form = CStr(Me.LabelTitle.Caption)
        FormCatchWizard.BtnImportOpenIssues.Enabled = True
        FormCatchWizard.BtnJustImport.Enabled = False
        FormCatchWizard.BtnSubmit.Enabled = False
        FormCatchWizard.BtnOsea.Enabled = False
        FormCatchWizard.Show vbModeless
    Else
        MsgBox "logika zatrzymana"
    End If
    
    ' ---------------------------------------------------------
    
        
    
End Sub

Private Sub BtnSubmit_Click()


    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
        inner_submit E_OI_CREATE
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
        inner_submit E_OI_UPDATE
    End If
    
End Sub

Private Sub inner_submit(e As E_OPEN_ISSUES_CRUD)
    
    ' text na guziki
    ' Global Const G_BTN_TEXT_ADD = "Dodaj"
    ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
    'Hide
    inner_calc e
    
    ' run_FormMain Me.LabelTitle
    Me.ListBox1.Clear
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    x = 1
    Dim r As Range
    Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Cells(1, 1)
    
    ' petla pobierania info z arkusza
    ' watpie ze kiedy kolwiek bardzie nagminne edytowania juz wpisanych issues zatem
    ' ogranicze sie do nazwania ich kolejnymi numerkami
    
    Do
        If CStr(Me.LabelTitle.Caption) = _
            CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then

        
                ' Exit Do
                Me.ListBox1.addItem "Open issue #" & CStr(x) & ", " & _
                    CStr(Trim(r.Offset(0, SIXP.e_open_issues_part_supplier - 1))) & ", " & _
                    r.Address
                x = x + 1
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    Me.ComboBoxStatus.Value = ""
    Me.TextBoxComment = ""
    Me.TextBoxDelivery = ""
    Me.TextBoxNoOfPNs = ""
    Me.TextBoxPartSupplier = ""
    Me.CheckBoxVisible.Value = True
End Sub

Private Sub change_col_M_in_MAIN_worksheet(ByRef r As Range)
    
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
                    rr.Offset(0, SIXP.e_main_last_update_on_open_issues - 1) = Trim(CStr(rr.Offset(0, 3)))
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

Private Sub inner_calc(e As E_OPEN_ISSUES_CRUD)


    'Public Enum E_XQ_ORDER
    '    e_xq_comment = 5
    '    e_xq_ppap_gate
    '    e_xq_project_type
    'End Enum


    
    If e = E_OI_CREATE Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = Trim(arr(x))
        Next x
        
        
        give_data_to_ranges r
        change_col_M_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf e = E_OI_UPDATE Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        v = Me.ListBox1.Value
        adr = adr_txt_parsed_from_selected_v_from_listbox(v)
    
        Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Range(CStr(adr))
        
        
        If CStr(Me.LabelTitle.Caption) = _
            CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
        
                give_data_to_ranges r
                change_col_M_in_MAIN_worksheet r
        End If
        
        
        ' ===================================================
    ElseIf e = E_OI_DELETE Then
        
        ' ===================================================
        ' przy delete no operation
        ' ===================================================
    Else
        MsgBox "fatal error on submitting!"
        End
    End If
End Sub

Private Sub give_data_to_ranges(ByRef r As Range)
    r.Parent.Cells(r.Row, SIXP.e_open_issues_comment) = CStr(Me.TextBoxComment)
    r.Parent.Cells(r.Row, SIXP.e_open_issues_delivery) = CStr(Me.TextBoxDelivery)
    r.Parent.Cells(r.Row, SIXP.e_open_issues_no_of_pn) = CStr(Me.TextBoxNoOfPNs)
    r.Parent.Cells(r.Row, SIXP.e_open_issues_part_supplier) = CStr(Me.TextBoxPartSupplier)
    r.Parent.Cells(r.Row, SIXP.e_open_issues_status) = CStr(Me.ComboBoxStatus)
    If CBool(Me.CheckBoxVisible) Then
        r.Parent.Cells(r.Row, SIXP.e_open_issues_visible) = "1"
    Else
        r.Parent.Cells(r.Row, SIXP.e_open_issues_visible) = "0"
    End If
End Sub



Private Sub ListBox1_Change()


    If Me.ListBox1.ListCount > 0 Then
    
        Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT


        ' Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Cells(1, 1)
        
        v = Me.ListBox1.Value
        adr = adr_txt_parsed_from_selected_v_from_listbox(v)
        
        Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Range(CStr(adr))
        
        If CStr(Me.LabelTitle) = _
            CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                ' -----------------------------------------------------------------------------------------------------------
                Me.TextBoxComment = CStr(r.Offset(0, SIXP.e_open_issues_comment - 1))
                Me.TextBoxDelivery = CStr(r.Offset(0, SIXP.e_open_issues_delivery - 1))
                Me.TextBoxNoOfPNs = CStr(r.Offset(0, SIXP.e_open_issues_no_of_pn - 1))
                Me.TextBoxPartSupplier = CStr(r.Offset(0, SIXP.e_open_issues_part_supplier - 1))
                Me.ComboBoxStatus = CStr(r.Offset(0, SIXP.e_open_issues_status - 1))
                If CStr(r.Offset(0, SIXP.e_open_issues_visible - 1)) = "1" Then
                    Me.CheckBoxVisible.Value = True
                Else
                    Me.CheckBoxVisible.Value = False
                End If
                ' -----------------------------------------------------------------------------------------------------------
        Else
            MsgBox "cos poszlo fest nie tak z pobraniem adresu komorki z open issues - program zatrzymal sie krytycznie"
            End
        End If
    Else
        Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD
    End If
End Sub

Private Function adr_txt_parsed_from_selected_v_from_listbox(v) As String
    
    gdzie_jest_pierwszy_dolar = Application.WorksheetFunction.Find("$", v)
    
    If gdzie_jest_pierwszy_dolar > 0 Then
        temp = Mid(v, gdzie_jest_pierwszy_dolar)
        
        temp = Replace(temp, "$", "")
        
        adr_txt_parsed_from_selected_v_from_listbox = temp
        
        ' w temp znajduje sie adres komorki ktora nas intererere
    Else
        ' cos poszlo nie tak
        MsgBox "cos poszlo fest nie tak z pobraniem adresu komorki z open issues - program zatrzymal sie krytycznie"
        End
    End If
End Function

' example
'Private Sub DTPickerPPAPGate_Change()''
'
' '   ' Me.TextBoxReleased = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerOrdersDue)))
'    Me.TextBoxPPAPGate = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerPPAPGate)))
'End Sub
