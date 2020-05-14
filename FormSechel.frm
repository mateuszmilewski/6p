VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSechel 
   Caption         =   "Sechel Handler"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14340
   OleObjectBlob   =   "FormSechel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormSechel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private schowek As String




Private Sub TextBoxAVenir_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxAVenir.Value)
    Me.LabelSchowek.Caption = schowek
End Sub


Private Sub TextBoxEnCours_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxEnCours.Value)
    Me.LabelSchowek.Caption = schowek
End Sub

Private Sub TextBoxFauxManquant_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxFauxManquant.Value)
    Me.LabelSchowek.Caption = schowek
End Sub

Private Sub TextBoxLines_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxLines.Value)
    Me.LabelSchowek.Caption = schowek
End Sub

Private Sub TextBoxManquant_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxManquant.Value)
    Me.LabelSchowek.Caption = schowek
End Sub

Private Sub TextBoxManquantPlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxManquantPlus.Value)
    Me.LabelSchowek.Caption = schowek
End Sub

Private Sub TextBoxRecu_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    schowek = CStr(Me.TextBoxRecu.Value)
    Me.LabelSchowek.Caption = schowek
End Sub




