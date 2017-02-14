Attribute VB_Name = "GetDataFromWizardBufferModule"
Public Function get_all_values(ptrn As String, r As Range) As Long

    get_all_values = 0
    
    
    
    Do
        If Trim(r) Like "*" & Trim(ptrn) & "*" Then
        
            If IsNumeric(r.Offset(1, 0)) Then
                get_all_values = get_all_values + CLng(r.Offset(1, 0))
            End If
        End If
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
    
    
    
    
End Function

Public Function get_value(ptrn As String, r As Range) As Long

    ' if its stay mean that we have some kind of issue
    get_value = -1
    
    Do
        If Trim(r) = Trim(ptrn) Then
        
            If IsNumeric(r.Offset(1, 0)) Then
                get_value = CLng(r.Offset(1, 0))
                Exit Do
            Else
                ' prawdopoodbnie "" jednak jak bedzie inny literal i tak zrzuce to na zero
                get_value = 0
                Exit Do
            End If
        End If
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
    
    
End Function

Public Function get_after_before_mrd(keyword, rl As Range, e As E_DEL_CONF_SPECIAL) As String
    get_after_before_mrd = ""
    
    lbl_str = Trim(Replace(CStr(ThisWorkbook.Sheets(SIXP.G_DEL_CONF_SPECIAL_SH_NM).Cells(e, 1)), "{MRD}", ""))
    lbl_str = Trim(Replace(CStr(lbl_str), "/", ""))
    
    If Trim(lbl_str) = "" Then
        lbl_str = "MRD"
    End If
    
    
    Dim tmp_r As Range
    Set tmp_r = rl
    
    Do
    
        If CStr(tmp_r) Like "*" & CStr(keyword) & "*" & CStr(lbl_str) & "*" Then
        
            get_after_before_mrd = CStr(tmp_r.Offset(1, 0))
            Exit Do
        End If
    
        Set tmp_r = tmp_r.Offset(0, 1)
    Loop Until Trim(tmp_r) = ""
End Function


Public Function get_del_conf_string_without_mrd(rl As Range, e As E_DEL_CONF_SPECIAL) As String
    get_del_conf_string_without_mrd = ""
    
    
    lbl_str = Trim(CStr(ThisWorkbook.Sheets(SIXP.G_DEL_CONF_SPECIAL_SH_NM).Cells(e, 1)))
    
    Dim tmp_r As Range
    Set tmp_r = rl
    
    Do
        If Trim(CStr(tmp_r)) = Trim(CStr(lbl_str)) Then
            get_del_conf_string_without_mrd = CStr(tmp_r.Offset(1, 0))
            Exit Do
        End If
    Loop Until Trim(r) = ""
End Function
