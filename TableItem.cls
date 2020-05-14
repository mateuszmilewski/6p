VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public arr As Collection
Public lines As Integer

Private Sub Class_Initialize()


    lines = 0
    Set arr = New Collection

End Sub

Private Sub Class_Terminate()
    Set arr = Nothing
End Sub
