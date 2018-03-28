VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOverseaDataCatcher 
   Caption         =   "Wybierz dane dla porcji Oversea..."
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6000
   OleObjectBlob   =   "FormOverseaDataCatcher.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOverseaDataCatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public activeSourceSheetName As String
Public labelka As String

Private Sub BtnSubmit_Click()


    If Me.BtnSubmit.Caption = "Submit" Then
        
        Hide
        ' MsgBox Me.TextBoxAdr.Value ' OK!
        
        Dim arr As Variant
        arr = Split(Me.TextBoxAdr.Value, ";")
        
        
        If LBound(arr) + 6 = UBound(arr) Then
        
            With FormOseaScope
                
                x = LBound(arr)
                .TextBoxTotal.Value = CStr(arr(x))
                .TextBoxOnStock.Value = CStr(arr(x + 1))
                .TextBoxOrdered.Value = CStr(arr(x + 2))
                .TextBoxConfirmed.Value = CStr(arr(x + 3))
                .TextBoxForMRD.Value = CStr(arr(x + 4))
                .TextBoxAfterMRD.Value = CStr(arr(x + 5))
                .TextBoxOPEN.Value = CStr(arr(x + 6))
            End With
        Else
            MsgBox "dane zostaly zle wybrane! krytyczne zamkniecie formularzy!"
            End
        End If
        
    ElseIf Me.BtnSubmit.Caption = "GET DATA" Then
        
        ' Me.TextBoxAdr.Value = activeSourceSheetName & "!" & Selection.Address
        Me.TextBoxAdr.Value = ""
        For Each r In Selection
            Me.TextBoxAdr.Value = Me.TextBoxAdr.Value & r & ";"
        Next r
        
        Me.TextBoxAdr.Value = Left(Me.TextBoxAdr.Value, Len(Me.TextBoxAdr.Value) - 1)
        
        Me.BtnSubmit.Caption = "Submit"
    
    End If
    
    
    
    
End Sub
