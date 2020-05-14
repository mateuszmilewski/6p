VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY Then
        Dim tbh As New TextBoxHandler
        tbh.showTextBoxByName Target, "TB_" & Replace(Target.Address, "$", ""), Target.Parent
    End If
End Sub
