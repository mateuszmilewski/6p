Attribute VB_Name = "BubbleSortModule"
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




Public Sub phase_list_sort(ByRef c() As Range)
    
    
    Dim pl As Worksheet, plr As Range
    Set pl = ThisWorkbook.Sheets(SIXP.G_PHASE_LIST_SH_NM)
    Set plr = pl.Range("A2")
    
    Set plr = pl.Range(plr, plr.End(xlDown))
    
    
    Dim orderForR1 As Integer, orderForR2 As Integer
    Dim r1 As Range, r2 As Range, o02r1 As Range, o02r2 As Range, tmp As Range
    
    
    For i = LBound(c) To UBound(c)
    
        If Not c(i + 1) Is Nothing Then
        
            
            Set r1 = c(i)
            Set r2 = c(i + 1)
            
            Set o02r1 = r1.Offset(0, 2)
            Set o02r2 = r2.Offset(0, 2)
            
            orderForR1 = Int(checkOrder(o02r1, plr))
            orderForR2 = Int(checkOrder(o02r2, plr))
            
        
        
            If Int(orderForR1) > Int(orderForR2) Then
            
                Set tmp = r1
                Set c(i) = r2
                Set c(i + 1) = r1
            
            
                ' zaczynamy od -1 poniewaz next i da interacje do przodu tak, czy inaczej
                i = LBound(c) - 1
            End If
        Else
            Exit For
        End If
    Next i
    
    
    
End Sub

Private Function checkOrder(r, plr) As Integer

    checkOrder = 100

    For Each iplr In plr
        
        If Trim(CStr(iplr.Offset(0, 1))) = Trim(CStr(r)) Then
            checkOrder = Int(iplr)
            Exit Function
        End If
    Next iplr
End Function

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
