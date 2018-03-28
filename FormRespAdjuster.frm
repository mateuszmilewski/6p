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
        
    
    
    SIXP.GlobalFooModule.gotoThisWorkbookMainA1
        
    Hide
    
    Dim rr As Range
    
    
    
    
    Set rr = ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")
    
    dopiszDoRejestruBrakujaceElemenciwa rr
    

    Dim znaleziony As Boolean
    znaleziony = False
    Do
        For x = 0 To Me.ListBoxInScope.ListCount - 1
        
    
                If Me.ListBoxInScope.List(x) = Trim(rr) Then
                    rr.Offset(0, 1) = 1
                    znaleziony = True
                    Exit For
                Else
                    
                    rr.Offset(0, 1) = 0
                End If
                

        Next x
        
        
        Set rr = rr.Offset(1, 0)
    Loop Until Trim(rr) = ""
    
    SIXP.LoadingFormModule.hideLoadingForm
    
End Sub


Private Sub dopiszDoRejestruBrakujaceElemenciwa(rr As Range)

    
    Dim obszarSzukania As Range
    
    
    Dim tmp As Range
    For x = 0 To Me.ListBoxInScope.ListCount - 1
    
        Set obszarSzukania = rr.Parent.Range(rr, rr.End(xlDown))
    
        Set tmp = Nothing
        On Error Resume Next
        Set tmp = obszarSzukania.Find(CStr(Me.ListBoxInScope.List(x)))
        
        
        If tmp Is Nothing Then
            rr.End(xlDown).Offset(1, 0) = CStr(Me.ListBoxInScope.List(x))
            rr.End(xlDown).Offset(1, 1) = 0
        End If
    Next x
    
End Sub

Private Sub ListBoxInScope_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


    Dim c1 As Collection
    'Dim c2 As Collection
    Dim item As String
    
    Set c1 = New Collection
    'Set c2 = New Collection
    
    

    For x = 0 To Me.ListBoxInScope.ListCount - 1
        
        If Me.ListBoxInScope.Selected(x) Then
            item = CStr(Me.ListBoxInScope.List(x))
        Else
            c1.Add CStr(Me.ListBoxInScope.List(x))
        End If
    Next x
    
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
    
    

    For x = 0 To Me.ListBoxOutOfScope.ListCount - 1
        
        If Me.ListBoxOutOfScope.Selected(x) Then
            item = CStr(Me.ListBoxOutOfScope.List(x))
        Else
            c1.Add CStr(Me.ListBoxOutOfScope.List(x))
        End If
    Next x
    
    Me.ListBoxOutOfScope.Clear
    
    For Each i In c1
        Me.ListBoxOutOfScope.AddItem CStr(i)
    Next i
    
    Me.ListBoxInScope.AddItem CStr(item)
End Sub

