VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadingForm 
   Caption         =   "Loading Form..."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   OleObjectBlob   =   "LoadingForm.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "LoadingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImageOff_Click()
    Me.ImageOff.Visible = False
    Me.ImageOn.Visible = True
    Hide
    
End Sub

Private Sub ImageOn_Click()
    Me.ImageOff.Visible = True
    Me.ImageOn.Visible = False
    Hide
    
End Sub


Private Sub Pasek1_Click()
    Hide
End Sub

Private Sub Pasek2Inner_Click()
    Hide
End Sub

Private Sub UserForm_Click()
    Hide
End Sub
