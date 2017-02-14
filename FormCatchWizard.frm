VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCatchWizard 
   Caption         =   "Catch Wizard"
   ClientHeight    =   4395
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

Private Sub BtnImportOpenIssues_Click()
    inner_ False
End Sub

Private Sub BtnJustImport_Click()
    inner_ False
End Sub

Private Sub BtnSubmit_Click()
    inner_ True
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' inner_ False
    MsgBox "no action under this click event!"
End Sub


Private Sub inner_(go_back As Boolean)


    Hide

    Dim wizard_workbook As Workbook
    
    Set wizard_workbook = Nothing
    
    On Error Resume Next
    Set wizard_workbook = Workbooks(Me.ListBox1.Value)
    
    
    
    Dim wh As WizardHandler
    Set wh = New WizardHandler
    
    wh.catch wizard_workbook
    
    
    If Me.czy_start_pochodzi_z_open_issues Then
    
        ' cala logika ktora opiera sie na pobieraniu danych z kolumny comments ktora staje sie zrodlem dla open issues
        ' -------------------------------------------------------------------------------------------------------------------
        '''
        ''
        Me.BtnImportOpenIssues.Enabled = True
        Me.BtnJustImport.Enabled = False
        Me.BtnSubmit.Enabled = False
        
        Dim m_sh As Worksheet, r As Range
        
        Dim oi As Worksheet
        Set oi = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
        
        Set m_sh = wizard_workbook.Sheets("MASTER")
        Set r = m_sh.Range("A2")
        
        ' -------------------------------------------------------------------------------------------------------------------
        ' pracujemy na wizard_workbook - musimy zaciagnac komentarze per duns!
        ''
        '
        ' przygotujemy najpierw slownik dla kazdego dunsu kolejne open issues
        '
        ''
        ' -------------------------------------------------------------------------------------------------------------------
        ''
        '''
        ' -------------------------------------------------------------------------------------------------------------------
    Else
        
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
                .CheckBoxWizardContent.Enabled = True
                .CheckBoxWizardContent.Value = False
                
                .Show
            End With
            
        End If
    End If
End Sub
