VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewDelConfEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public adapterArray As Collection
Private ea As EventAdapter

Private Sub Class_Initialize()


    Set adapterArray = New Collection


    Dim x As Integer
    For x = 1 To 128
        Set ea = New EventAdapter
        adapterArray.Add ea
    Next x
End Sub
