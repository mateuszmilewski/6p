Attribute VB_Name = "QuikModule"
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
