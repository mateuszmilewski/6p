VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCatchWizard 
   Caption         =   "Catch Wizard"
   ClientHeight    =   3180
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

Private Sub BtnSubmit_Click()
    inner_
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    inner_
End Sub


Private Sub inner_()


    Hide

    Dim wizard_workbook As Workbook
    
    Set wizard_workbook = Nothing
    
    On Error Resume Next
    Set wizard_workbook = Workbooks(Me.ListBox1.Value)
    
    
    
    Dim wh As WizardHandler
    Set wh = New WizardHandler
    
    wh.catch wizard_workbook
    
    With NewProj
        .TextBoxCW = wh.get_cw()
        .TextBoxFaza = wh.get_faza()
        .TextBoxPlt = wh.get_plt()
        .TextBoxProj = wh.get_proj()
        .ComboBoxStatus = SIXP.GlobalCrossTriangleCircleModule.putCross
        .Show
    End With
End Sub
