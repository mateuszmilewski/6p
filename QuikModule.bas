Attribute VB_Name = "QuikModule"
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

Enum TypQuika
    STD
    arr
End Enum



Public Sub quicksort(typ_quika As TypQuika, ByRef r As Range, Optional p, Optional k)

    If IsMissing(p) And IsMissing(k) Then
        p = 1
        k = r.Count
    End If
    
    If p < k Then
        q = podziel(typ_quika, r, p, k)
        quicksort typ_quika, r, p, q - 1
        quicksort typ_quika, r, q + 1, k
    End If
End Sub

Private Function podziel(typ_quika As TypQuika, ByRef r As Range, p, k)
    
    gold = r.item(k)
    b = p
    
    For X = p To k
        If gold >= r.item(X) Then
            If typ_quika = STD Then
                zamien_ze_soba r.item(b), r.item(X)
            ElseIf typ_quika = arr Then
                zamien_ze_soba_arr r.item(b), r.item(X)
            End If
            b = b + 1
        End If
    Next X
    
    podziel = b - 1
End Function

Private Sub zamien_ze_soba_arr(ByRef a As Range, ByRef b As Range)
                wiersz_b = a.Row
                wiersz_x = b.Row
                
                Dim rng As Range
                Set rng = ActiveSheet.UsedRange
                fst_col = rng.item(1).Column
                lst_col = rng.item(rng.Count).Column
                
                zamien_ze_soba Range(Cells(wiersz_b, fst_col), Cells(wiersz_b, lst_col)), _
                    Range(Cells(wiersz_x, fst_col), Cells(wiersz_x, lst_col))
End Sub

Private Sub zamien_ze_soba(ByRef a As Range, ByRef b As Range)
    Dim tmp As Variant
    
    For X = 1 To a.Count
        tmp = a.item(X).Value
        a.item(X).Value = b.item(X).Value
        b.item(X).Value = tmp
    Next X
End Sub

Public Sub main_szybki_sort()
    quicksort STD, Selection
End Sub

Public Sub main_szybki_sort_arr()
    quicksort arr, Selection
End Sub
