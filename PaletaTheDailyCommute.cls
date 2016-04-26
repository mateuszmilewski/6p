VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PaletaTheDailyCommute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dark_grey As Long
Public medium_grey As Long
Public light_grey As Long

Public yellow As Long
Public orange As Long


Private Sub Class_Initialize()

    dark_grey = RGB(61, 60, 58)
    medium_grey = RGB(112, 112, 112)
    light_grey = RGB(221, 212, 188)
    yellow = RGB(248, 220, 11)
    orange = RGB(251, 168, 39)

End Sub

Private Sub Class_Terminate()

End Sub
