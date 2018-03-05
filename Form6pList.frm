VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form6pList 
   Caption         =   "6P files"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9390
   OleObjectBlob   =   "Form6pList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form6pList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnClose_Click()
    Hide
End Sub

Private Sub BtnOpen_Click()


    SIXP.GlobalFooModule.gotoThisWorkbookMainA1

    Hide
    
    
    
    answer = MsgBox("Pliki nie spelniajace standardu zostana pominiete... czy chcesz kontynuowac?", vbYesNo)
    
    If answer = vbYes Then
        ' and now action!
        ' --------------------------------------------------------
        ' --------------------------------------------------------
        
        
        
        Dim wrk As Workbook
        Dim filename As String
        For x = 0 To Me.ListBoxIn.ListCount - 1
            filename = CStr(Me.ListBoxIn.List(x))
            
            innerRunLogicFor6P2 filename
        Next x
        
        
        
        MsgBox "ready!"
        
        ' --------------------------------------------------------
        ' --------------------------------------------------------
    End If
End Sub

Private Sub ListBoxIn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


    For x = 0 To Me.ListBoxIn.ListCount - 1
        If Me.ListBoxIn.Selected(x) Then
            Me.ListBoxOut.AddItem Me.ListBoxIn.List(x)
            Me.ListBoxIn.RemoveItem x
        End If
    Next x
    
End Sub

Private Sub ListBoxOut_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


    For x = 0 To Me.ListBoxOut.ListCount - 1
        If Me.ListBoxOut.Selected(x) Then
            Me.ListBoxIn.AddItem Me.ListBoxOut.List(x)
            Me.ListBoxOut.RemoveItem x
        End If
    Next x
End Sub
