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

    Dim rc As RecentBuildPlanChangesHandler
    Set rc = New RecentBuildPlanChangesHandler
    
    With SIXP.FormRecentBuildPlanChanges
        rc.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormRecentBuildPlanChanges
        .Show
    End With
End Sub

Public Sub zrob_contracted_pnoc(edit_czy_add As String, txt_from_combo_box As String)

    Dim cp As ContractedPNOCHandler
    Set cp = New ContractedPNOCHandler
    
    With SIXP.FormContractedPNOC
        cp.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormContractedPNOC
        .Show
    End With
    
End Sub

Public Sub zrob_osea_scope(edit_czy_add As String, txt_from_combo_box As String)

    Dim osea As OseaScopeHandler
    Set osea = New OseaScopeHandler
    
    With SIXP.FormOseaScope
        osea.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormOseaScope
        .Show
    End With
End Sub


Public Sub zrob_total(edit_czy_add As String, txt_from_combo_box As String)

    Dim tot As Total5PHandler
    Set tot = New Total5PHandler
    
    With SIXP.FormTotals5p
        'Application.EnableEvents = False
        tot.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormTotals5p
        .Show
        'Application.EnableEvents = True
    End With
End Sub

Public Sub zrob_xq(edit_czy_add As String, txt_from_combo_box As String)
    
    Dim xq As XQ6PHandler
    Set xq = New XQ6PHandler
    
    With SIXP.FormX6P
        xq.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormX6P
        .Show
    End With
    

End Sub

Public Sub zrob_del_conf(edit_czy_add As String, txt_from_combo_box As String)

    Dim dc As DelConfStatus7XHandler
    Set dc = New DelConfStatus7XHandler
    
    With SIXP.FormDelConfStatus
        dc.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormDelConfStatus
        .Show
    End With
End Sub

Public Sub zrob_open_issues(edit_czy_add As String, txt_from_combo_box As String)

    Dim zoi As OpenIssues8XHandler
    Set zoi = New OpenIssues8XHandler
    
    With SIXP.FormOpenIssues
        zoi.label = txt_from_combo_box
        zoi.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormOpenIssues
        .Show
    End With
End Sub

Public Sub zrob_resp(edit_czy_add As String, txt_from_combo_box As String)

    Dim resp As Resp9XHandler
    Set resp = New Resp9XHandler
    
    With SIXP.FormResp
        resp.inicjuj_wstepne_wartosci_pol_do_wypelnienia CStr(txt_from_combo_box), CStr(edit_czy_add), SIXP.FormResp
        .Show
    End With

End Sub
