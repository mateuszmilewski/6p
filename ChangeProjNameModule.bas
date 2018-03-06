Attribute VB_Name = "ChangeProjNameModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.



Public Sub change_project_name(ictrl As IRibbonControl)


    If SIXP.GlobalFooModule.checkIfFirstFourFieldsProjektPlantCodeFazaCW(ActiveSheet) Then
        If ActiveCell.Row > 1 Then
            If ActiveCell.Parent.Cells(ActiveCell.Row, 1).Value <> "" Then
                openModelessFormAccordinglyTo ActiveCell
            Else
                MsgBox "Puste dane!"
            End If
        Else
            MsgBox "nie mozesz wybrac pierwszego wiersza!"
        End If
    End If
    
End Sub


Private Sub openModelessFormAccordinglyTo(r As Range)


    Dim l As T_Link
    Set l = New T_Link
    
    l.zrob_mnie_z_range r.Parent.Cells(r.Row, 1)
    
    przygotujFormularzZmianyNazwyProjektu r.Parent.Cells(r.Row, 1), l

End Sub

Private Sub przygotujFormularzZmianyNazwyProjektu(r As Range, l As T_Link)
    
    
    With SIXP.FormChangeProjectName
        .TextBoxCurrProj.Value = l.project
        .TextBoxCurrPltCode.Value = l.plt
        .TextBoxCurrFaza.Value = l.faza
        .TextBoxCurrCw.Value = l.cw
        
        .Show vbModeless
    End With
End Sub
