VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public d As Dictionary
Public maxLines As Integer


Private ti As TableItem


Public Sub addItem(item As String, klucz)

    Set ti = New TableItem
    
    splitted = Split(item, chr(10))
    
    ti.lines = UBound(splitted) + 1
    
    If ti.lines > maxLines Then
        maxLines = ti.lines
    End If
    
    For x = LBound(splitted) To UBound(splitted)
        ti.arr.Add splitted(x)
    Next x
    
    d.Add klucz, ti
    
End Sub



Private Sub Class_Initialize()
    maxLines = 1
    Set d = New Dictionary
End Sub

Private Sub Class_Terminate()
    maxLines = 0
    Set d = Nothing
End Sub
