VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOrderReleaseStatus 
   Caption         =   "Order Release Status"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   OleObjectBlob   =   "FormOrderReleaseStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOrderReleaseStatus"
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



Private Sub BtnCopy_Click()

    copyOneItemFromDifferentRecord Me.name
End Sub

Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnSubmit_Click()


    Set walidator = New Validator
    With walidator
        .dodajDoKolekcji Me.TextBoxMRD, .pStr_checkIfYYYYCW
        .dodajDoKolekcji Me.TextBoxBuild, .pStr_checkIfYYYYCW
        .dodajDoKolekcji Me.TextBoxBOMFreeze, .pStr_checkIfYYYYCW
        
        .dodajDoKolekcji Me.TextBoxNoOfVeh, .pStr_checkIfNumber
        
        .dodajDoKolekcji Me.TextBoxOrdersDue, .pStr_checkIfYYYYCW
        .dodajDoKolekcji Me.TextBoxReleased, .pStr_checkIfYYYYCW
        
        .dodajDoKolekcji Me.TextBoxWeeksDelay, .pStr_checkIfNumber
        
        .run
    End With
    

    If walidator.pass Then

        SIXP.GlobalFooModule.gotoThisWorkbookMainA1
    
        ' text na guziki
        ' Global Const G_BTN_TEXT_ADD = "Dodaj"
        ' Global Const G_BTN_TEXT_EDIT = "Edytuj"
        'Hide
        inner_calc
        
        'run_FormMain Me.LabelTitle
        
        If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
            Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT
        End If
    Else
        MsgBox "Walidation failed!"
    End If
End Sub

Private Sub change_col_F_in_MAIN_worksheet(ByRef r As Range)
    
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
                    rr.Offset(0, SIXP.e_main_last_update_on_order_release_status - 1) = Trim(CStr(rr.Offset(0, 3)))
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
        Set r = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = Trim(arr(x))
        Next x
        
        
        give_data_to_ranges r
        change_col_F_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm).Cells(1, 1)
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_F_in_MAIN_worksheet r
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
    r.Parent.Cells(r.Row, SIXP.e_order_release_mrd) = CStr(Me.TextBoxMRD)
    r.Parent.Cells(r.Row, SIXP.e_order_release_build) = CStr(Me.TextBoxBuild)
    r.Parent.Cells(r.Row, SIXP.e_order_release_bom_freeze) = CStr(Me.TextBoxBOMFreeze)
    r.Parent.Cells(r.Row, SIXP.e_order_release_no_of_veh) = CStr(Me.TextBoxNoOfVeh)
    r.Parent.Cells(r.Row, SIXP.e_order_release_orders_due) = CStr(Me.TextBoxOrdersDue)
    r.Parent.Cells(r.Row, SIXP.e_order_release_released) = CStr(Me.TextBoxReleased)
    r.Parent.Cells(r.Row, SIXP.e_order_release_weeks_delay) = CStr(Me.TextBoxWeeksDelay)
End Sub




' template from NewProj
'

' DTPICKERS!
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------
'Private Sub ComboBoxPLT_Change()
'    Me.TextBoxPlt = CStr(Me.ComboBoxPLT.Value)
'End Sub
'
'Private Sub DTPicker1_Change()
'    Me.TextBoxCW = SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPicker1.Value))
'End Sub

Private Sub DTPickerMRD_Change()
    Me.TextBoxMRD = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerMRD.Value)))
End Sub

Private Sub DTPickerOrdersDue_Change()
    Me.TextBoxOrdersDue = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerOrdersDue)))
End Sub

Private Sub DTPickerReleased_Change()
    Me.TextBoxReleased = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerOrdersDue)))
End Sub

Private Sub DTPickerBuild_Change()
    Me.TextBoxBuild = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerBuild)))
End Sub

Private Sub DTPickerBOMFreeze_Change()
    Me.TextBoxBOMFreeze = CStr(SIXP.GlobalFooModule.parse_from_date_to_yyyycw(CDate(Me.DTPickerBOMFreeze)))
End Sub
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------

' textboxes with qtyies
' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------

Private Sub NoOfVehLess_Click()
    If IsNumeric(Me.TextBoxNoOfVeh) Then
        If CLng(Me.TextBoxNoOfVeh) > 0 Then
            tmp = CLng(Me.TextBoxNoOfVeh)
            tmp = tmp - 1
            Me.TextBoxNoOfVeh = CStr(tmp)
        End If
    End If
End Sub

Private Sub NoOfVehLess_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxNoOfVeh) Then
        If CLng(Me.TextBoxNoOfVeh) > 9 Then
            tmp = CLng(Me.TextBoxNoOfVeh)
            tmp = tmp - 10
            Me.TextBoxNoOfVeh = CStr(tmp)
        End If
    End If
End Sub

Private Sub NoOfVehMore_Click()
    If IsNumeric(Me.TextBoxNoOfVeh) Then
        tmp = CLng(Me.TextBoxNoOfVeh)
        tmp = tmp + 1
        Me.TextBoxNoOfVeh = CStr(tmp)
    End If
End Sub


Private Sub NoOfVehMore_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.TextBoxNoOfVeh) Then
        tmp = CLng(Me.TextBoxNoOfVeh)
        tmp = tmp + 10
        Me.TextBoxNoOfVeh = CStr(tmp)
    End If
End Sub

Private Sub TryWizardBtn_Click()
    
    ' sub odpowiadajacy za sciaganie danych z wizard buff worksheet
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    '3: MRD
    '4: BUILD START
    '5: BUILD END
    '6: BOM
    '7: PPAP GATE
    
    With buff
        
        Me.TextBoxBOMFreeze.Value = Replace(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Cells(1, 6)), "CW", ""), "Y", "")
        Me.TextBoxBuild.Value = Replace(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Cells(1, 4)), "CW", ""), "Y", "")
        Me.TextBoxMRD.Value = Replace(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Cells(1, 3)), "CW", ""), "Y", "")
        Me.TextBoxOrdersDue = ""
        Me.TextBoxReleased = ""
        Me.TextBoxNoOfVeh = 0
        Me.TextBoxWeeksDelay = 0
    End With
End Sub

Private Sub WeeksDelayLess_Click()
     If IsNumeric(Me.TextBoxWeeksDelay) Then
        
        If CLng(Me.TextBoxWeeksDelay) > 0 Then
            tmp = CLng(Me.TextBoxWeeksDelay)
            tmp = tmp - 1
            Me.TextBoxWeeksDelay = CStr(tmp)
        End If
     End If
End Sub

Private Sub WeeksDelayMore_Click()
    If IsNumeric(Me.TextBoxWeeksDelay) Then
        tmp = CLng(Me.TextBoxWeeksDelay)
        tmp = tmp + 1
        Me.TextBoxWeeksDelay = CStr(tmp)
    End If
End Sub

' ------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------



