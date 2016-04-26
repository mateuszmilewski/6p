Attribute VB_Name = "FormMainModule"
Public Sub run_FormMain()
    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.init
    
    Set fmh = Nothing
End Sub


