Attribute VB_Name = "LoadingFormModule"
Option Explicit

Private Declare Function FindWindow Lib "user32" _
Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function DrawMenuBar Lib "user32" ( _
ByVal hwnd As Long) As Long

Private Sub RemoveTitleBar(frm As Object)
    Dim lStyle          As Long
    Dim hMenu           As Long
    Dim mhWndForm       As Long
     
    If val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", frm.Caption) 'for Office 97 version
    Else
        mhWndForm = FindWindow("ThunderDFrame", frm.Caption) 'for office 2000 or above
    End If
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
End Sub


Public Sub showLoadingForm()
    
    SIXP.LoadingFormModule.RemoveTitleBar SIXP.LoadingForm
    
    
    initLoadingPage
    
    With SIXP.LoadingForm
        .Height = 150
        .Top = Application.Top + Application.Height * 0.5 - .Height * 0.5
        .Left = Application.Left + Application.Width * 0.5 - .Width * 0.5
        .Show vbModeless
    End With
    
    
End Sub


Public Sub hideLoadingForm()

    SIXP.LoadingForm.Hide
    
End Sub

Private Sub initLoadingPage()
    With SIXP.LoadingForm
        .Pasek2Inner.Width = 1
        .ImageOn.Visible = True
        .ImageOff.Visible = False
    End With
End Sub

Public Sub incLoadingForm()
    
    increaseLoadingFormStatus 10
End Sub

Public Sub increaseLoadingFormStatus(dx As Integer)
    increasePasekCustom Int(dx)
End Sub

Private Sub increasePasekCustom(dx As Integer)
    With SIXP.LoadingForm
        .Pasek2Inner.Width = .Pasek2Inner.Width + Int(dx)
        
        If .Pasek2Inner.Width > .Pasek1.Width Then
            .Pasek2Inner.Width = 1
        End If
        
        
        If .Pasek2Inner.Width > 100 And .Pasek2Inner.Width < 150 Then
            toggleVisiblityBetweenImagesInLoadingForm
        End If
        
        
        If .Pasek2Inner.Width > 170 And .Pasek2Inner.Width < 200 Then
            toggleVisiblityBetweenImagesInLoadingForm
        End If
        
        If .Pasek2Inner.Width > 200 Then
            .ImageOn.Visible = True
            .ImageOff.Visible = False
        End If
        
        
        .Repaint
        
    End With
End Sub

Private Sub toggleVisiblityBetweenImagesInLoadingForm()
    
    With SIXP.LoadingForm
        If .ImageOff.Visible Then
            .ImageOn.Visible = True
            .ImageOff.Visible = False
        Else
            .ImageOff.Visible = True
            .ImageOn.Visible = False
        End If
    End With
End Sub

Private Sub testOnLoadingPage()
    
    showLoadingForm
    
    Dim x As Integer
    For x = 1 To 500
        increaseLoadingFormStatus 10
        Sleep 100
    Next x
    
    SIXP.LoadingForm.Hide
    
End Sub


