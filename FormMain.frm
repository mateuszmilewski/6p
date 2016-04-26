VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMain 
   Caption         =   "Main Interface"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   OleObjectBlob   =   "FormMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnOrderReleaseStatus_Click()
    Hide
    zrob_order_release_status CStr(Me.BtnOrderReleaseStatus.Caption), CStr(Me.ComboBoxProject.Value)
End Sub
