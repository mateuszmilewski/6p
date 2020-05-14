Attribute VB_Name = "ClearModule"
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

Public Sub clear_item(ictrl As IRibbonControl)
    innerclearitem
End Sub

Public Sub advanced_clearing(ictrl As IRibbonControl)
    
    AdvClearForm.Show vbModeless
End Sub


Private Sub innerclearitem()
    
    Dim s As SixPChecker
    Set s = New SixPChecker
    
    If s.sprawdz_czy_arkusz_aktywny_to_ten_arkusz Then
        ' to jest ogolenie sprawdzenie czy pierwsze 4 kolumny to kolumny std proj, plt, faza, cw
        If s.sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main Then
            
            Dim dm As DeletionManager
            Set dm = New DeletionManager
            
            
            dm.usun_kazde_wystapienie_dla_aktywnej_komorki ActiveCell
            
            Set dm = Nothing
        End If
    End If
End Sub



' pod formularz advanced clearing
Public Sub clear_all_items()
    Dim dm As DeletionManager
    Set dm = New DeletionManager
    SIXP.LoadingFormModule.showLoadingForm
    SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    dm.usun_wszystko
    SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    SIXP.LoadingFormModule.hideLoadingForm
    Set dm = Nothing
End Sub

Public Sub clear_by_wildcard(pattern As String)
    Dim dm As DeletionManager
    Set dm = New DeletionManager
    SIXP.LoadingFormModule.showLoadingForm
    SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    dm.usun_po_gwiazdce pattern
    SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    SIXP.LoadingFormModule.hideLoadingForm
    Set dm = Nothing
End Sub


Public Sub close_excel_project_reports(ictrl As IRibbonControl)


    Application.ScreenUpdating = False

    SIXP.LoadingFormModule.showLoadingForm

    Dim w As Workbook
    For Each w In Workbooks
        If w.name Like "*" & CStr(SIXP.G_EXCEL_REPORT_PREFIX) & "*" Then
            SIXP.LoadingFormModule.incLoadingForm
            w.Close False
            
        End If
    Next w
    
    SIXP.LoadingFormModule.hideLoadingForm
    
    
    Application.ScreenUpdating = True
    
    
    MsgBox "ready!"
End Sub



Public Sub remove_one_item_from_one_sheet(ictrl As IRibbonControl)



    Dim l As T_Link
    Dim lr As Linker
    Dim e As E_MAIN_ORDER
    
    Set l = New T_Link
    
    
    
    If SIXP.GlobalFooModule.checkIfFirstFourFieldsProjektPlantCodeFazaCW(ActiveSheet) Then
        
        If CStr(ActiveSheet.name) = CStr(SIXP.G_main_sh_nm) Then
        
            
        
            If ActiveCell.Column > 5 And ActiveCell.Row > 1 And ActiveCell.Value <> "" Then
            
                
                ' body
                ' -------------------------------------------------------------------------------------------------------
                l.zrob_mnie_z_range ActiveCell.Parent.Cells(ActiveCell.Row, 1)
                
                e = usunWierszSideowy(ActiveCell, l)  ' ActiveCell.Parent.Cells(ActiveCell.Row, 1)
                usunWpisWMain ActiveCell, True, l, e
                
                ' -------------------------------------------------------------------------------------------------------
            Else
                MsgBox "nie mozesz dla tej selekcji nic usunac!"
            End If
        Else
        
            If ActiveCell.Row > 1 And ActiveCell.Value <> "" Then
            
                ' body
                ' -------------------------------------------------------------------------------------------------------
                l.zrob_mnie_z_range ActiveCell.Parent.Cells(ActiveCell.Row, 1)
                
                e = usunWierszSideowy(ActiveCell, l)  ' ActiveCell.Parent.Cells(ActiveCell.Row, 1)
                usunWpisWMain ActiveCell, False, l, e
                
                ' -------------------------------------------------------------------------------------------------------
            Else
                MsgBox "nie mozesz dla tej selekcji nic usunac!"
            End If
        
        End If
    End If
End Sub


Private Function usunWierszSideowy(r As Range, l As T_Link) As E_MAIN_ORDER

    Dim side As Worksheet
    
    If r.Parent.name <> CStr(SIXP.G_main_sh_nm) Then
    
        Set side = r.Parent
        usunWierszSideowy = znajdzKolumne(CStr(side.name))
    Else
    
        Set side = znajdzSidea(Int(r.Column))
        usunWierszSideowy = Int(r.Column)
    End If
    
    
    Dim elDoUsuniecia As Range
    Set elDoUsuniecia = l.znajdz_siebie_w_arkuszu(side)
    
    elDoUsuniecia.EntireRow.Delete xlShiftUp
    
    

End Function


Private Function znajdzSidea(ktoraKolumna As Integer) As Worksheet

    Set znajdzSidea = Nothing
    
    Dim e As E_MAIN_ORDER
    e = ktoraKolumna
    
    If e = e_main_last_update_on_order_release_status Then
    
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
    
    ElseIf e = e_main_last_update_on_recent_build_plan_changes Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
        
    ElseIf e = e_main_last_update_on_chart_contracted_pnoc Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
        
    ElseIf e = e_main_last_update_on_osea Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
        
    ElseIf e = e_main_last_update_on_totals Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
        
    ElseIf e = e_main_last_update_on_xq Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
        
    ElseIf e = e_main_last_update_on_del_conf Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
        
    ElseIf e = e_main_last_update_on_open_issues Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
        
    ElseIf e = e_main_last_update_on_resp Then
    
        Set znajdzSidea = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
        
    End If
    
End Function


Private Function znajdzKolumne(nazwaArkuszaSideowego As String) As E_MAIN_ORDER

    
    Dim nm As String
    nm = CStr(nazwaArkuszaSideowego)
    
    If nm = CStr(SIXP.G_order_release_status_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_order_release_status
    ElseIf nm = CStr(SIXP.G_recent_build_plan_changes_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_recent_build_plan_changes
    ElseIf nm = CStr(SIXP.G_cont_pnoc_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_chart_contracted_pnoc
    ElseIf nm = CStr(SIXP.G_osea_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_osea
    ElseIf nm = CStr(SIXP.G_totals_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_totals
    ElseIf nm = CStr(SIXP.G_xq_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_xq
    ElseIf nm = CStr(SIXP.G_del_conf_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_del_conf
    ElseIf nm = CStr(SIXP.G_open_issues_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_open_issues
    ElseIf nm = CStr(SIXP.G_resp_sh_nm) Then
        znajdzKolumne = e_main_last_update_on_resp
    End If
End Function

    
Private Sub usunWpisWMain(r As Range, inMain As Boolean, l As T_Link, e As E_MAIN_ORDER)


    If inMain Then
        
        r.Value = ""
    Else
    
        Dim sh As Worksheet
        Set sh = ThisWorkbook.Sheets(CStr(SIXP.G_main_sh_nm))
        
        Set r = r.Parent.Cells(r.Row, 1)
        
        Dim tmp As Range
        Set tmp = l.znajdz_siebie_w_arkuszu(sh)
        
        If Not tmp Is Nothing Then
        
            tmp.Parent.Cells(tmp.Row, e).Value = ""
        Else
            MsgBox "Arkusz MAIN nie posiada tego wpisu, zatem zostanie on usuniety tylko w tym arkuszu!"
        End If
    End If
    
End Sub




Public Sub clearNewTableSheet()
    
    Dim newTableSh As Worksheet
    Set newTbaleSh = ThisWorkbook.Sheets("newTable")
    
    Dim r As Range
    Set r = newTbaleSh.Range("A1:A548576")

    'Selection.Delete Shift:=xlUp
    r.EntireRow.Delete xlShiftUp

End Sub

Public Sub checkIfGreyAndRemoveWithShiftUp(r As Range)

    Dim newTableSh As Worksheet
    Set newTbaleSh = ThisWorkbook.Sheets("newTable")
    

    If r.Interior.Color <> RGB(200, 200, 200) Then
        r.EntireRow.Delete xlShiftUp
    End If
End Sub


Public Sub clearNewOnePagerWorksheet()


    Dim noprsh As Worksheet
    Set noprsh = ThisWorkbook.Sheets("NEW ONE PAGER")
    
    Dim r As Range
    Set r = noprsh.Range("A500:A548576")
    r.EntireRow.Delete xlShiftUp

End Sub
