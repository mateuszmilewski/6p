Attribute VB_Name = "DataFlowPodFormMainModule"
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

' niniejszy modul (DataFlowPodFormMainModule) zawiera wszystkie niezbedne suby,
' ktore uruchamiaja sie po kliknieciu w przyciski formularza
' FormMain

' sekcja subow, ktore pracuja po kliknieciu w odpowiednie guziki
Public Sub zrob_order_release_status(edit_czy_add As String, txt_from_combo_box As String)

    Dim o As OrderReleaseStatusHandler
    Set o = New OrderReleaseStatusHandler
    
    ' ten modul nie zawiera jeszcze implementacji do uruchomienia formularza
    ' de facto obiekt powyzej tez ma pusta implementacje jeszcze
    
    
    With FormOrderReleaseStatus

        o.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), FormOrderReleaseStatus
        
        .Show
    End With
    
    Set o = Nothing

End Sub

Public Sub zrob_recent_build_plan_changes(edit_czy_add As String, txt_from_combo_box As String)

End Sub
