Attribute VB_Name = "BubbleSortModule"
' bubble_range_sort
Public Sub bubble_sort(ByRef c() As Range, ofst, e As E_SORT_TYPE)
    
    Dim tmp As Range, r1 As Range, r2 As Range
    
    Dim ii As Long, iplus1 As Long, wynik As Long
    
    If e = E_DESCENDING Then
    
    
        For i = LBound(c) To UBound(c)
        
            If Not c(i + 1) Is Nothing Then
            
            
                Set r1 = c(i)
                Set r2 = c(i + 1)
                
                ii = r1.Offset(0, ofst)
                iplus1 = r2.Offset(0, ofst)
                wynik = CLng(ii) - CLng(iplus1)
                If wynik < 0 Then
                
                    Set tmp = r1
                    Set c(i) = r2
                    Set c(i + 1) = r1
                
                    i = LBound(c) - 1
                End If
            Else
                Exit For
            End If
        
        Next i
    ElseIf e = E_ASCENDING Then
    
    
        For i = 0 To UBound(c)
        
           If Not c(i + 1) Is Nothing Then
            
            
                Set r1 = c(i)
                Set r2 = c(i + 1)
                
                ii = r1.Offset(0, ofst)
                iplus1 = r2.Offset(0, ofst)
                wynik = CLng(iplus1) - CLng(ii)
                If wynik < 0 Then
                
                    Set tmp = r1
                    Set r1 = r2
                    Set r2 = r1
                
                    i = -1
                End If
            Else
                Exit For
            End If
        
        Next i
    End If
End Sub



' std bubble sort on array with values only
Public Sub bubble_sort_on_values(ByRef c(), ofst, e As E_SORT_TYPE)
    
    If e = E_DESCENDING Then
    
    
        For i = LBound(c) To UBound(c)
        
            If Not c(i + 1) Is Nothing Then
            
            
                v1 = c(i)
                v2 = c(i + 1)
                
                If IsNumeric(v1) And IsNumeric(v2) Then
                    wynik = CLng(v1) - CLng(v2)
                Else
                
                    casted1 = asciien(CStr(v1))
                    casted2 = asciien(CStr(v2))
                    
                    wynik = CLng(casted1) - CLng(casted2)
                End If
                
                If wynik < 0 Then
                
                    tmp = v1
                    c(i) = v2
                    c(i + 1) = v1
                
                    i = LBound(c) - 1
                End If
            Else
                Exit For
            End If
        
        Next i
    ElseIf e = E_ASCENDING Then
    
    
        For i = 0 To UBound(c)
        
           If Not c(i + 1) Is Nothing Then
            
                v1 = c(i)
                v2 = c(i + 1)
                
                If IsNumeric(v1) And IsNumeric(v2) Then
                    wynik = CLng(v2) - CLng(v1)
                Else
                
                    casted1 = asciien(CStr(v1))
                    casted2 = asciien(CStr(v2))
                    
                    wynik = CLng(casted2) - CLng(casted1)
                End If
                
                
                If wynik < 0 Then
                
                    tmp = v1
                    c(i) = v2
                    c(i + 1) = v1
                
                    i = -1
                End If
            Else
                Exit For
            End If
        
        Next i
    End If
End Sub


Private Function asciien(s As String) As String
' Returns the string to its respective ascii numbers
   Dim i As Integer

   For i = 1 To Len(s)
      asciien = asciien & CStr(Asc(Mid(s, i, 1)))
   Next i

End Function
