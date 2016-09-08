VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function return_full_concated_r_string(link As T_Link) As String
    
    return_concated_r_string = CStr(link.project) & CStr(link.plt) & CStr(link.faza) & CStr(link.cw)
End Function

Public Function return_full_concated_r_string_comma_seperated(link As T_Link) As String
    
    return_concated_r_string = CStr(link.project) & ", " & CStr(link.plt) & ", " & CStr(link.faza) & ", " & CStr(link.cw)
End Function

Public Function return_full_concated_r_string_without_cw(link As T_Link) As String
    
    return_full_concated_r_string_without_cw = CStr(link.project) & CStr(link.plt) & CStr(link.faza)
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
        For x = 1 To Len(slowo)
            liczba = liczba & CStr(Asc(Mid(slowo, x, 1)))
        Next x
    End If
    
    convert_all_word_into_no = liczba
End Function




