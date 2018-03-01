VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
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




Public Function return_full_concated_r_string(link As T_Link) As String
    
    tmp = Trim(link.project) & Trim(link.plt) & Trim(link.faza) & Trim(link.cw)
    tmp = Replace(tmp, " ", "")
    tmp = Replace(tmp, ",", "")
    tmp = Replace(tmp, ".", "")
    tmp = Replace(tmp, "/", "")
    tmp = Replace(tmp, "\", "")
    
    
    return_full_concated_r_string = CStr(tmp)
End Function

Public Function return_full_concated_r_string_comma_seperated(link As T_Link) As String
    
    return_full_concated_r_string_comma_seperated = Trim(link.project) & ", " & Trim(link.plt) & ", " & Trim(link.faza) & ", " & Trim(link.cw)
End Function

Public Function return_full_concated_r_string_without_cw(link As T_Link) As String
    
    return_full_concated_r_string_without_cw = Trim(link.project) & Trim(link.plt) & Trim(link.faza)
End Function


Public Function retrun_short_number(link As T_Link) As String
    ' init
    retrun_short_number = CLng(0)
    
    
    ' now calc
    retrun_short_number = convert_all_word_into_no(link.project) & _
        convert_all_word_into_no(link.plt) & _
        convert_all_word_into_no(link.faza) & _
        convert_all_word_into_no(link.cw)
    
End Function

Public Function retrun_short_number_without_cw(link As T_Link) As String
    ' init
    retrun_short_number_without_cw = CLng(0)
    
    
    ' now calc
    retrun_short_number_without_cw = convert_all_word_into_no(link.project) & _
        convert_all_word_into_no(link.plt) & _
        convert_all_word_into_no(link.faza)
    
End Function

Private Function convert_all_word_into_no(slowo As String) As String

    liczba = ""

    If Len(slowo) > 0 Then
        For X = 1 To Len(slowo)
            liczba = liczba & CStr(Asc(Mid(slowo, X, 1)))
        Next X
    End If
    
    convert_all_word_into_no = liczba
End Function




