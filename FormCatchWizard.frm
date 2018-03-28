VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCatchWizard 
   Caption         =   "Catch Wizard"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormCatchWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCatchWizard"
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


Public czy_start_pochodzi_z_open_issues As Boolean
Public wez_text_z_open_issues_form As String

Public labelTextForOverseaPortion As String

Private Sub BtnGetFrom6P_Click()
    
    SIXP.GlobalFooModule.gotoThisWorkbookMainA1
    
    Hide
    
    Me.czy_start_pochodzi_z_open_issues = False

    ' in import from 6p module
    ' -------------------------------------
    innerRunLogicFor6P Me
    ' -------------------------------------
    MsgBox "import from 6P ready!"
End Sub

Private Sub BtnImportOpenIssues_Click()

    Me.czy_start_pochodzi_z_open_issues = True

    inner_ False
End Sub

Private Sub BtnJustImport_Click()

    Me.czy_start_pochodzi_z_open_issues = False
    inner_ False
End Sub

Private Sub BtnOsea_Click()
    Hide
    Me.czy_start_pochodzi_z_open_issues = False
    externalOverseaDataCatchLogic Me.ListBox1.Value, Me, Me.labelTextForOverseaPortion
End Sub

Private Sub BtnSubmit_Click()
    Me.czy_start_pochodzi_z_open_issues = False
    inner_ True
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' inner_ False
    Me.czy_start_pochodzi_z_open_issues = False
    MsgBox "no action under this click event!"
End Sub


Private Sub inner_(go_back As Boolean)



    SIXP.GlobalFooModule.gotoThisWorkbookMainA1

    Hide
    
    SIXP.LoadingFormModule.showLoadingForm
    
    Application.ScreenUpdating = False

    Dim wizard_workbook As Workbook
    
    Set wizard_workbook = Nothing
    
    On Error Resume Next
    Set wizard_workbook = Workbooks(Me.ListBox1.Value)
    
    SIXP.LoadingFormModule.increaseLoadingFormStatus 20
    
    Dim wh As WizardHandler
    Set wh = New WizardHandler
    
    wh.catch wizard_workbook
    
    
    If Me.czy_start_pochodzi_z_open_issues Then
    
        ' cala logika ktora opiera sie na pobieraniu danych z kolumny comments ktora staje sie zrodlem dla open issues
        ' -------------------------------------------------------------------------------------------------------------------
        '''
        ''
        'Me.BtnImportOpenIssues.Enabled = True
        'Me.BtnJustImport.Enabled = False
        'Me.BtnSubmit.Enabled = False
        
        
        ' jednak zanim cos zrobimy sprawdzmy jeszcze czy mamy zaciagniety label
        If Me.wez_text_z_open_issues_form <> "" Then
        
        
            
            
        
            Dim m_sh As Worksheet, r As Range, details_sh As Worksheet
            
            Dim oi As Worksheet
            Set oi = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
            
            Set m_sh = wizard_workbook.Sheets("MASTER")
            Set details_sh = wizard_workbook.Sheets("DETAILS")
            Set r = m_sh.Range("A2")
            
            SIXP.LoadingFormModule.incLoadingForm
            
            
            ' ale nawet jesli mamy juz tekst z open issues fajnie by bylo sprawdzic czy to jest ten sam td
            ' If sprawdz_zgodnosc_danych_miedzy_arkuszem_master_a_labelka(wh) Then
            If True Then ' usunalem to srogie sprawdzanie danych ponowne!
            
                ' -------------------------------------------------------------------------------------------------------------------
                ' pracujemy na wizard_workbook - musimy zaciagnac komentarze per duns!
                ''
                '
                ' przygotujemy najpierw slownik dla kazdego dunsu kolejne open issues
                Dim oih As OpenIssues8XHandler
                Set oih = New OpenIssues8XHandler
                
                SIXP.LoadingFormModule.incLoadingForm
                SIXP.LoadingFormModule.hideLoadingForm
                
                oih.currProj = CStr(wez_text_z_open_issues_form)
                
                With oih
                   ' .label = Me.wez_text_z_open_issues_formd
                    If .wypelnij_slownik_dunsami_jako_kluczami_pod_open_issues_z_mass_impotu_z_wizarda(oi, m_sh, r, details_sh) Then
                        .zrzuc_dane_ze_slownika_do_arkusza_open_issues
                        .usun_duplikaty
                        .otworz_formularz_ponownie
                    Else
                        .otworz_formularz_ponownie_bez_danych Me.wez_text_z_open_issues_form
                    End If
                End With
            Else
                MsgBox "niezgodnosc danych miedzy rekordem z ktorego uruchomiono mass import a wizardem!"
            End If
        End If
        '
        ''
        ' -------------------------------------------------------------------------------------------------------------------
        ''
        '''
        ' -------------------------------------------------------------------------------------------------------------------
    Else
    
        wh.go_with_6p_time
        
        SIXP.LoadingFormModule.incLoadingForm
        
        Me.BtnImportOpenIssues.Enabled = False
        Me.BtnJustImport.Enabled = True
        Me.BtnSubmit.Enabled = True
        
        If go_back Then
            
            With NewProj
                .TextBoxCW = wh.get_cw()
                .TextBoxFaza = wh.get_faza()
                .TextBoxPlt = wh.get_plt()
                .TextBoxProj = wh.get_proj() & " " & wh.get_biw_ga() & " " & " MY: " & wh.get_my
                .ComboBoxStatus = SIXP.GlobalCrossTriangleCircleModule.putCross
                
                
                
                ' zmiana od wersji 0.26
                ' 2017-01-24
                ' -----------------------------------------------
                '
                ' odblokowuje opcje zaciagania z buffa ale musi jeszcze user dokliknac
                ' bylo blokowanie, ale jednak zostawie to - dam mu zaciagnac dane normalnie
                ' walidacje dam nieco pozniej czy te dane sie trzymaja kupy
                ' .CheckBoxWizardContent.Enabled = True
                .CheckBoxWizardContent.Value = False
                
                SIXP.LoadingFormModule.hideLoadingForm
                
                .Show vbModeless
            End With
            
        End If
    End If
End Sub


Private Function sprawdz_zgodnosc_danych_miedzy_arkuszem_master_a_labelka(wh As WizardHandler) As Boolean
    sprawdz_zgodnosc_danych_miedzy_arkuszem_master_a_labelka = False
    
    
    ' TODO: check compatibility between data
    ' ---------------------------------------------------------------------------------------
    
    ' sprawdzmy czy wystepuje ta sama proj, faza, plt
    
    ' tutaj na proj mniejsza restrykcja: zamiast sprawdzac czy faktycznie sa spacje miedzy kolejnymi argumentami
    ' wrzucam po prostu wildcardy
    If CStr(Me.wez_text_z_open_issues_form) Like "*" & wh.get_proj() & "*MY*" & wh.get_my & "*" Then
        If CStr(Me.wez_text_z_open_issues_form) Like "*" & CStr(wh.get_faza) & "*" Then
            If CStr(Me.wez_text_z_open_issues_form) Like "*" & CStr(wh.get_plt) & "*" Then
            
                ' dosyc ostre warunki zostaly spelnione
                ' mozemy mysle zabrac sie za import danych
                ' -------------------------------------------------------------------------------------
                sprawdz_zgodnosc_danych_miedzy_arkuszem_master_a_labelka = True
                ' -------------------------------------------------------------------------------------
            End If
        End If
    End If
    
    ' ---------------------------------------------------------------------------------------
End Function
