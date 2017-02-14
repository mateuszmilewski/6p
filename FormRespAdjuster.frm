VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormRespAdjuster 
   Caption         =   "Resp Adjuster"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8160
   OleObjectBlob   =   "FormRespAdjuster.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormRespAdjuster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnFMALeft_Click()


    Dim rr As Range
    
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")
    
    Do
        If CStr(rr) Like "*FMA*" Then
            rr.Offset(0, 1) = 1
        Else
            rr.Offset(0, 1) = 0
        End If
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""
    
    
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")
    Me.ListBoxInScope.Clear
    Me.ListBoxOutOfScope.Clear
    
    Do
        If CStr(rr) Like "*FMA*" Then
            ' rr.Offset(0, 1) = 1
            Me.ListBoxInScope.AddItem CStr(rr)
        Else
            ' rr.Offset(0, 1) = 0
            Me.ListBoxOutOfScope.AddItem CStr(rr)
        End If
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""
    
    
End Sub

Private Sub BtnSubmit_Click()
    Hide
    
    Dim rr As Range
    
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")

    Do
        For X = 0 To Me.ListBoxInScope.ListCount - 1
        
    
                If Me.ListBoxInScope.List(X) = Trim(rr) Then
                    rr.Offset(0, 1) = 1
                    Exit For
                Else
                    
                    rr.Offset(0, 1) = 0
                End If
                

        Next X
        
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""
    
End Sub

Private Sub ListBoxInScope_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


    Dim c1 As Collection
    'Dim c2 As Collection
    Dim item As String
    
    Set c1 = New Collection
    'Set c2 = New Collection
    
    

    For X = 0 To Me.ListBoxInScope.ListCount - 1
        
        If Me.ListBoxInScope.Selected(X) Then
            item = CStr(Me.ListBoxInScope.List(X))
        Else
            c1.Add CStr(Me.ListBoxInScope.List(X))
        End If
    Next X
    
    Me.ListBoxInScope.Clear
    
    For Each i In c1
        Me.ListBoxInScope.AddItem CStr(i)
    Next i
    
    Me.ListBoxOutOfScope.AddItem CStr(item)

End Sub

Private Sub ListBoxOutOfScope_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    
    Dim c1 As Collection
    'Dim c2 As Collection
    Dim item As String
    
    Set c1 = New Collection
    'Set c2 = New Collection
    
    

    For X = 0 To Me.ListBoxOutOfScope.ListCount - 1
        
        If Me.ListBoxOutOfScope.Selected(X) Then
            item = CStr(Me.ListBoxOutOfScope.List(X))
        Else
            c1.Add CStr(Me.ListBoxOutOfScope.List(X))
        End If
    Next X
    
    Me.ListBoxOutOfScope.Clear
    
    For Each i In c1
        Me.ListBoxOutOfScope.AddItem CStr(i)
    Next i
    
    Me.ListBoxInScope.AddItem CStr(item)
End Sub
