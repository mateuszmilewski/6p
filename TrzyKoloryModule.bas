Attribute VB_Name = "TrzyKoloryModule"
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
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

Public Sub reset_3_kolory_form()
    
    
    Dim c As Control
    
    With Form3Kolory
        For Each c In .Controls
        
            If TypeName(c) = "TextBox" Then

                c.Value = 0
            End If
        Next c
        
        fill_combobox
        
        ' .Show
        .Repaint
    End With
End Sub


Public Sub fill_it()

    ' MsgBox "not implemented yet!"

    Dim c As Control
    
    With Form3Kolory
        For Each c In .Controls
        
            If TypeName(c) = "TextBox" Then
                If c.Value = "" Then
                    c.Value = 0
                End If
            End If
        Next c
        
        fill_combobox
        
        .Show vbModeless
    End With
    
    
End Sub



Private Sub fill_combobox()


    Form3Kolory.ComboBoxLink.Clear


    Dim l As T_Link, lr As Linker
    Set lr = New Linker
    Dim r As Range, del_conf_sh As Worksheet
    Set del_conf_sh = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
    
    
    Set r = del_conf_sh.Cells(2, 1)
    Do
        Set l = New T_Link
        l.zrob_mnie_z_range r
        txt = CStr(lr.return_full_concated_r_string_comma_seperated(l))
        
        Form3Kolory.ComboBoxLink.addItem txt
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""

    Form3Kolory.ComboBoxLink.Value = ""
    
End Sub
