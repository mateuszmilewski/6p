VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Public WithEvents tbx As MSForms.TextBox
Attribute tbx.VB_VarHelpID = -1


Private Sub tbx_Change()


    ' OK
    ' Debug.Print "tbx changed! - " & tbx.name & ", " & tbx.Value
    If GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY Then
    
        If Trim(tbx.Value) <> "" Then
    
            If tbx.name Like "TextBoxG*" Then recalc_sums "G"
            If tbx.name Like "TextBoxY*" Then recalc_sums "Y"
            If tbx.name Like "TextBoxR*" Then recalc_sums "R"
        End If
    End If
End Sub


Private Sub recalc_sums(chr As String)
    
    With SIXP.FormDelConfStatus2
        
        If chr = "G" Then
            .TextBoxGSUM.Value = 0
            On Error Resume Next
            .TextBoxGSUM.Value = sumujGreeny()
        End If
        
        If chr = "Y" Then
            .TextBoxYSUM.Value = 0
            On Error Resume Next
            .TextBoxYSUM.Value = sumujYellowy()
        End If
        
        If chr = "R" Then
            .TextBoxRSUM.Value = 0
            On Error Resume Next
            .TextBoxRSUM.Value = sumujRedy()
        End If
        
        
        
        
        
        .TextBoxTotalFromBoxes.Value = _
            CLng(.TextBoxGSUM.Value) + CLng(.TextBoxYSUM.Value) + CLng(.TextBoxRSUM.Value)
    End With
    
End Sub


Private Function sumujGreeny() As Long
    sumujGreeny = 0
    
    On Error Resume Next
    sumujGreeny = innerSum("G")
End Function

Private Function sumujYellowy() As Long
    sumujYellowy = 0
    
    On Error Resume Next
    sumujYellowy = innerSum("Y")
End Function

Private Function sumujRedy() As Long
    sumujRedy = 0
    
    On Error Resume Next
    sumujRedy = innerSum("R")
End Function


Private Function innerSum(chr As String) As Long
    
    
    innerSum = 0
    
    Dim c As MSForms.Control, tbx As MSForms.TextBox
    
    For Each c In SIXP.FormDelConfStatus2.Controls
    
        If (c.name Like "TextBox" & CStr(chr) & "*") And (Not c.name = "TextBox" & CStr(chr) & "SUM") Then
            
            Set tbx = c
            If IsNumeric(tbx.Value) Then innerSum = innerSum + CLng(tbx.Value)
        End If
    Next c
End Function
