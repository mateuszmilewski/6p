Attribute VB_Name = "FormMainModule"
Public Sub run_FormMain()
    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.init
    
    Set fmh = Nothing
End Sub

Public Sub add_new_project()
    
    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.new_project
    
    Set fmh = Nothing
End Sub


