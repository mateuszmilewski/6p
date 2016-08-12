Attribute VB_Name = "GlobalFooModule"
Public Function parse_from_date_to_yyyycw(d As Date) As String
    
    
    cstr_y = CStr(Year(d))
    cstr_iso_cw = CStr(Application.WorksheetFunction.IsoWeekNum(CDbl(d)))
    
    If Len(cstr_iso_cw) = 2 Then
        parse_from_date_to_yyyycw = cstr_y & cstr_iso_cw
    ElseIf Len(cstr_iso_cw) = 1 Then
        parse_from_date_to_yyyycw = cstr_y & "0" & cstr_iso_cw
    End If
End Function


Public Function from_yyyy_cw_to_monday_from_this_week(yyyycw As String) As Date
    
    y = Left(yyyycw, 4)
    cw = Right(yyyycw, 2)
        
    ' -------------------- ' -------------------- ' --------------------
    
    Dim d As Date
    d = CDate(y & "-01-01")
    
    Do
        d = d + 1
    Loop Until CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(d))) = CLng(cw)
    
    from_yyyy_cw_to_monday_from_this_week = d
End Function
