VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDelConfStatus2 
   Caption         =   "Delivery Confirmation Status"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8865
   OleObjectBlob   =   "FormDelConfStatus2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormDelConfStatus2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
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


Private eh As NewDelConfEventHandler
Private ea As EventAdapter


Private walidator As Validator



Private Sub BtnGoBack_Click()
    Hide
    run_FormMain Me.LabelTitle
End Sub

Private Sub BtnSubmit_Click()
    
    Dim c As Control
    Set walidator = New Validator
    With walidator
        
        For Each c In Me.Controls

            ' not really like this solution...
            ' top 3 textbox have name with prefix Special!
            If c.name Like "TextBox*" Then
                
                ' we have inside tb real textbox object!
                ' -------------------------------------------
                Debug.Print c.name & " ready for validation!"
                
                .dodajDoKolekcji c, .pStr_checkIfNumber
                
                ' -------------------------------------------
            End If
            
            
        Next c
        
        
        
        .run
    End With
    
    
    If walidator.pass Then
    
        SIXP.GlobalFooModule.gotoThisWorkbookMainA1
    
    
        new_inner_calc
    
        If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
            Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT
        End If
    Else
        MsgBox "Validation failed!"
    End If
End Sub


Private Sub new_inner_calc()



    Dim r As Range
    
    If Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_ADD Then
    
        ' no to szukamy pierwszego wolnego pola i wsadzamy
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
        Do
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
        
        
        Dim arr As Variant
        arr = Split(CStr(Me.LabelTitle), ",")
        For x = 0 To 3
            r.Offset(0, x) = Trim(arr(x))
        Next x
        
        
        give_data_to_ranges r
        change_col_L_in_MAIN_worksheet r
        
        ' tutaj raczej bledu wychwytywac nie bedziemy - chodzi o zwyczajne (z pewnoscia)
        ' dodanie info na sam koniec tabeli
        
        
        
        ' ===================================================
    
    ElseIf Me.BtnSubmit.Caption = SIXP.G_BTN_TEXT_EDIT Then
    
    
        ' szukamy jeszcze raz
        ' ===================================================
        Set r = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Cells(1, 1)
        Do
            If CStr(Me.LabelTitle.Caption) = _
                CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
            
                    give_data_to_ranges r
                    change_col_L_in_MAIN_worksheet r
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

    ' ---------------------------------------------------------------------------------------
    With r.Parent
        .Cells(r.Row, SIXP.e_del_conf_after_alt) = CStr(Me.TextBoxYALT)
        .Cells(r.Row, SIXP.e_del_conf_after_alttwomrd) = CStr(Me.TextBoxYALTTWO)
        .Cells(r.Row, SIXP.e_del_conf_after_mrd) = CStr(Me.TextBoxYMRD)
        .Cells(r.Row, SIXP.e_del_conf_after_oncostmrd) = CStr(Me.TextBoxYONCOST)
        .Cells(r.Row, SIXP.e_del_conf_after_salttwomrd) = CStr(Me.TextBoxYSALTTWO)
        .Cells(r.Row, SIXP.e_del_conf_after_smrd) = CStr(Me.TextBoxYSMRD)
        .Cells(r.Row, SIXP.e_del_conf_after_soncostmrd) = CStr(Me.TextBoxYSONCOST)
        ' .cells(r.Row, sixp.e_del_conf_after_twomrd) = "" ' obsoelete
        
        .Cells(r.Row, SIXP.e_del_conf_edi) = CStr(Me.TextBoxGEDI)
        .Cells(r.Row, SIXP.e_del_conf_for_alt) = CStr(Me.TextBoxGALT)
        .Cells(r.Row, SIXP.e_del_conf_for_alttwomrd) = CStr(Me.TextBoxGALTTWO)
        .Cells(r.Row, SIXP.e_del_conf_for_mrd) = CStr(Me.TextBoxGMRD)
        .Cells(r.Row, SIXP.e_del_conf_for_oncostmrd) = CStr(Me.TextBoxGONCOST)
        .Cells(r.Row, SIXP.e_del_conf_for_salttwomrd) = CStr(Me.TextBoxGSALTTWO)
        .Cells(r.Row, SIXP.e_del_conf_for_smrd) = CStr(Me.TextBoxGSMRD)
        .Cells(r.Row, SIXP.e_del_conf_for_soncostmrd) = CStr(Me.TextBoxGSONCOST)
        ' .cells(r.Row, sixp.e_del_conf_for_twomrd) = "" ' obsolete
        
        
        .Cells(r.Row, SIXP.e_del_conf_ho) = CStr(Me.TextBoxGHO)
        .Cells(r.Row, SIXP.e_del_conf_na) = CStr(Me.TextBoxGNA)
        .Cells(r.Row, SIXP.e_del_conf_on_stock) = CStr(Me.TextBoxGONSTOCK)
        .Cells(r.Row, SIXP.e_del_conf_open) = CStr(Me.TextBoxROpen)
        .Cells(r.Row, SIXP.e_del_conf_pot_itdc) = CStr(Me.TextBoxRPOTITDC)
        
        .Cells(r.Row, SIXP.e_del_conf_red_after_alt) = CStr(Me.TextBoxRALT)
        .Cells(r.Row, SIXP.e_del_conf_red_after_alttwomrd) = CStr(Me.TextBoxRALTTWO)
        .Cells(r.Row, SIXP.e_del_conf_red_after_mrd) = CStr(Me.TextBoxRMRD)
        .Cells(r.Row, SIXP.e_del_conf_red_after_oncostmrd) = CStr(Me.TextBoxRONCOST)
        .Cells(r.Row, SIXP.e_del_conf_red_after_salttwomrd) = CStr(Me.TextBoxRSALTTWO)
        .Cells(r.Row, SIXP.e_del_conf_red_after_smrd) = CStr(Me.TextBoxRSMRD)
        .Cells(r.Row, SIXP.e_del_conf_red_after_soncostmrd) = CStr(Me.TextBoxRSONCOST)
        ' .cells(r.Row, sixp.e_del_conf_too_late) = "" ' obsolete
        .Cells(r.Row, SIXP.e_del_conf_yellow_open) = CStr(Me.TextBoxYOpen)
        
        
        
        
        
        
    End With
    ' ---------------------------------------------------------------------------------------
End Sub


Private Sub change_col_L_in_MAIN_worksheet(ByRef r As Range)
    

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
                    rr.Offset(0, SIXP.e_main_last_update_on_del_conf - 1) = Trim(CStr(rr.Offset(0, 3)))
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




Private Sub DelConfFromBuffBtn_Click()

    
    GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = False
    Application.EnableEvents = False
    
    
    Dim buff As Worksheet
    Set buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    
    ' rngv - range values, range labels
    Dim rngv As Range, rngl As Range, h1_total
    Set rngv = buff.Cells(17, 1)
    Set rngl = buff.Cells(16, 1)
    Set h1_total = buff.Range("H1")
    
    
    Dim m_r_d As Range, build_start As Range
    
    Set m_r_d = buff.Range("C1")
    Set build_start = buff.Range("D1")
    
    ' -------------------------------------------------------------------------------------
    ''
    '
    
    
    'zeros
    Me.TextBoxROpen = 0
    Me.TextBoxYOpen = 0
    Me.TextBoxRPOTITDC = 0

    
    
    With buff
    
        ' 15: BEFORE OR ON/AFTER MRD
        ' 16: BEFORE/AFTER MRD - labels all
        ' 17: values
        'Me.TextBoxAfterALTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_MRD))
        'Me.TextBoxForALTMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_MRD))
        
        
        
        
        
        ' GREEN szajs
        ' ----------------------------------------------------------------------------------------------------------------------------------
        Me.TextBoxGALT = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_MRD))
        Me.TextBoxGALTTWO = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ALT_TWO_MRD))
        Me.TextBoxGMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_MRD))
        Me.TextBoxGONCOST = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_ONCOST_MRD))
        Me.TextBoxGSALTTWO = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        Me.TextBoxGSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_MRD))
        Me.TextBoxGSONCOST = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("BEFORE", rngl, E_DCS_Staggered_ONCOST_MRD))
        ' ----------------------------------------------------------------------------------------------------------------------------------
        
        
        ' YELLOW szajs
        ' ----------------------------------------------------------------------------------------------------------------------------------
        Me.TextBoxYALT = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_MRD))
        Me.TextBoxYALTTWO = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ALT_TWO_MRD))
        Me.TextBoxYMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_MRD))
        Me.TextBoxYONCOST = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_ONCOST_MRD))
        Me.TextBoxYSALTTWO = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        Me.TextBoxYSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_MRD))
        Me.TextBoxYSONCOST = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER", rngl, E_DCS_Staggered_ONCOST_MRD))
        ' ----------------------------------------------------------------------------------------------------------------------------------
        
        
        ' RED szajs
        ' ----------------------------------------------------------------------------------------------------------------------------------
        Me.TextBoxRALT = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_ALT_MRD))
        Me.TextBoxRALTTWO = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_ALT_TWO_MRD))
        Me.TextBoxRMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_MRD))
        Me.TextBoxRONCOST = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_ONCOST_MRD))
        Me.TextBoxRSALTTWO = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_Staggered_ALT_TWO_MRD))
        Me.TextBoxRSMRD = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_Staggered_MRD))
        Me.TextBoxRSONCOST = CStr(SIXP.GetDataFromWizardBufferModule.get_after_before_mrd("AFTER BUILD START", rngl, E_DCS_Staggered_ONCOST_MRD))
        ' ----------------------------------------------------------------------------------------------------------------------------------
    End With
        
        
        
    
    ' DEL CONF, WHICH IS NOT CONNECTED WITH MRD PARAM.
    
    Set rngv = buff.Cells(13, 1)
    Set rngl = buff.Cells(12, 1)
    
    
    Application.EnableEvents = False
    
    With buff
        
        Me.TextBoxGEDI = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_EDI))
        Me.TextBoxGHO = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_HO))
        Me.TextBoxGNA = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_NA))
        Me.TextBoxGONSTOCK = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_ON_STOCK))
        
        
        Me.TextBoxRPOTITDC = CStr(SIXP.GetDataFromWizardBufferModule.get_del_conf_string_without_mrd(rngl, E_DCS_POTITDC))
    End With
    
    
    ' trzeba teraz podliczyc sumy oraz open fields!
    
    Set rngv = buff.Range(buff.Cells(17, 1), buff.Cells(17, 100))
    suma_17 = Application.WorksheetFunction.Sum(rngv)
    Set rngv = buff.Range(buff.Cells(13, 1), buff.Cells(13, 100))
    suma_13 = Application.WorksheetFunction.Sum(rngv)
    
    suma_pod_open = 0
    suma_nie_open = suma_17 + suma_13
    
    suma_all = 0
    If IsNumeric(buff.Range("h1").Value) Then suma_all = CLng(buff.Range("h1").Value)
    
    If suma_all > 0 Then
        suma_pod_open = suma_all - suma_nie_open
    End If

    If sprawdz_czy_open_jest_zolte_czy_moze_jest_czerwone(buff) = "RED" Then
        Me.TextBoxROpen.Value = suma_pod_open
        Me.TextBoxYOpen.Value = 0
    ElseIf sprawdz_czy_open_jest_zolte_czy_moze_jest_czerwone(buff) = "YELLOW" Then
        Me.TextBoxROpen.Value = 0
        Me.TextBoxYOpen.Value = suma_pod_open
    End If
    
    '
    ''
    ' -------------------------------------------------------------------------------------
    
    
    Application.EnableEvents = True
    GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = True
End Sub


Private Function sprawdz_czy_open_jest_zolte_czy_moze_jest_czerwone(b As Worksheet) As String
    sprawdz_czy_open_jest_zolte_czy_moze_jest_czerwone = "RED"
    
    
    
    Dim yyyycwBOM As String
    yyyycwBOM = Replace(Replace(CStr(b.Range("F1").Value), "CW", ""), "Y", "")
    
    Dim BOMDate As Date
    BOMDate = SIXP.GlobalFooModule.from_yyyy_cw_to_monday_from_this_week(yyyycwBOM) ' to jest sroga restryckja => musimy to przenisc na koniec tygodnia
    
    ' proste przeliczenie do soboty :)
    BOMDate = BOMDate + 6
    
    If CDate(Date) <= CDate(BOMDate) Then
        sprawdz_czy_open_jest_zolte_czy_moze_jest_czerwone = "YELLOW"
    End If
    
    
End Function




Private Sub UserForm_Deactivate()
    Set eh = Nothing
    Set ea = Nothing
End Sub

Private Sub UserForm_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Set eh = Nothing
    Set ea = Nothing
End Sub

Private Sub UserForm_Initialize()
    
    If eh Is Nothing Then
        
        Set eh = New NewDelConfEventHandler
        
        
        Dim tbx As Object
        Dim ctbx As MSForms.TextBox
        
        i = 1
        For Each tbx In Me.Controls
        
            Set ctbx = Nothing
        
            If tbx.name Like "TextBoxG*" Or tbx.name Like "TextBoxY*" Or tbx.name Like "TextBoxR*" Then
                tbx.Value = 0
            
                If Not tbx.name Like "*SUM*" Then
                    Set ctbx = tbx
                    Set ea = eh.adapterArray.item(i)
                    Set ea.tbx = ctbx
                    i = i + 1
                End If
            End If
            
        Next tbx
        
        
        ' take data from H1
        TextBoxTotalFromH1.Value = CStr(ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("h1").Value)
    
    End If
End Sub

Private Sub UserForm_Terminate()
    Set eh = Nothing
    Set ea = Nothing
End Sub
