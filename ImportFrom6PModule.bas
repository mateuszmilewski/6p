Attribute VB_Name = "ImportFrom6PModule"
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


Public Sub import_from_another_6p(ictrl As IRibbonControl)



    

    
    With FormCatchWizard
    
        .ListBox1.Clear
        .ListBox1.MultiSelect = fmMultiSelectMulti
        
        .BtnGetFrom6P.Enabled = True
        .BtnImportOpenIssues.Enabled = False
        .BtnJustImport.Enabled = False
        .BtnSubmit.Enabled = False
        
        
        
    End With
        
    For Each w In Workbooks
        With FormCatchWizard.ListBox1
            If CStr(w.name) <> CStr(ThisWorkbook.name) Then
                .AddItem CStr(w.name)
            End If
        End With
    Next w
    
    FormCatchWizard.Show vbModeless

End Sub

Public Sub innerRunLogicFor6P2(filename As String)



    SIXP.LoadingFormModule.showLoadingForm



    Application.ScreenUpdating = False

    Dim wrk As Workbook
    ' Set wrk = Workbooks(CStr(frm.ListBox1.Value)) ' when mutli available you can not just simple take value
    
    
    Dim wrkCollection As Collection
    Set wrkCollection = New Collection
    
    
    SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    

    Set wrk = Nothing
    Set wrk = Workbooks(CStr(filename))
    wrkCollection.Add wrk
    SIXP.LoadingFormModule.incLoadingForm

    
    For Each wrk In wrkCollection
        
        SIXP.LoadingFormModule.increaseLoadingFormStatus 100
        If checkIfYouCanMigrateData(wrk) Then
            migrateDataBetween wrk, ThisWorkbook ' private subs from leanData already in...
        Else
            'MsgBox "wybrany plik: " & CStr(wrk.FullName) & " nie spelnia standardow!", vbCritical
            'End
        End If
    Next wrk
    
    
    Application.ScreenUpdating = True
    
    SIXP.LoadingFormModule.hideLoadingForm
End Sub


Public Sub innerRunLogicFor6P(frm As FormCatchWizard)



    SIXP.LoadingFormModule.showLoadingForm



    Application.ScreenUpdating = False

    Dim wrk As Workbook
    ' Set wrk = Workbooks(CStr(frm.ListBox1.Value)) ' when mutli available you can not just simple take value
    
    
    Dim wrkCollection As Collection
    Set wrkCollection = New Collection
    
    
    SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    
    For x = 0 To frm.ListBox1.ListCount - 1
        If frm.ListBox1.Selected(x) Then
            Set wrk = Nothing
            Set wrk = Workbooks(CStr(frm.ListBox1.List(x)))
            wrkCollection.Add wrk
            SIXP.LoadingFormModule.incLoadingForm
        End If
    Next x
    
    For Each wrk In wrkCollection
        
        SIXP.LoadingFormModule.increaseLoadingFormStatus 100
        If checkIfYouCanMigrateData(wrk) Then
            migrateDataBetween wrk, ThisWorkbook ' private subs from leanData already in...
        Else
            MsgBox "wybrany plik: " & CStr(wrk.FullName) & " nie spelnia standardow!", vbCritical
            End
        End If
    Next wrk
    
    
    Application.ScreenUpdating = True
    
    SIXP.LoadingFormModule.hideLoadingForm
End Sub


Public Sub make_6p_lean(ictrl As IRibbonControl)
    leanData
End Sub


Private Sub leanData()
    
    ' ten sub bedzie dzialal po wykananiu zaciagniecia danych
    ' tj. usunie entire row jesli dane sa puste pomiedzy
    ' plus sprawdzi ktore z wierszy uzupelnione sa tylko i wylacznie pierwszymi czterema kolumnami
    ' =====================================================================================================
    ' =====================================================================================================
    
    SIXP.LoadingFormModule.showLoadingForm
    
    Application.ScreenUpdating = False
    
    Dim c As Collection
    Set c = New Collection
    
    ' troche manualnie ale ostatecznie podjalem decyzje, ze niech to tak zostanie, przynajmniej jest czytelne
    ' i widac wyraznie co, gdzie i jak :D
    c.Add SIXP.G_main_sh_nm
    c.Add SIXP.G_order_release_status_sh_nm
    c.Add SIXP.G_recent_build_plan_changes_sh_nm
    c.Add SIXP.G_cont_pnoc_sh_nm
    c.Add SIXP.G_osea_sh_nm
    c.Add SIXP.G_totals_sh_nm
    c.Add SIXP.G_resp_sh_nm
    c.Add SIXP.G_del_conf_sh_nm
    c.Add SIXP.G_open_issues_sh_nm
    c.Add SIXP.G_xq_sh_nm
    
    SIXP.LoadingFormModule.incLoadingForm
    
    
    
    Dim destSh As Worksheet
    For Each shnm In c
    
        Set destSh = ThisWorkbook.Sheets(shnm)
        
        
        SIXP.LoadingFormModule.increaseLoadingFormStatus 100
    
        removeDuplicatesInDest destSh
        SIXP.LoadingFormModule.incLoadingForm
        removeEmptyBetween destSh
        SIXP.LoadingFormModule.incLoadingForm
        doSomeLeanFor destSh
        SIXP.LoadingFormModule.incLoadingForm
        entireRowsRemovalIfEmpty destSh
        SIXP.LoadingFormModule.incLoadingForm
    Next shnm
    
    SIXP.LoadingFormModule.incLoadingForm
    
    ' =====================================================================================================
    ' =====================================================================================================
    
    Application.ScreenUpdating = False
    
    SIXP.LoadingFormModule.hideLoadingForm
    
    MsgBox "ready!"
    
    
End Sub


Private Function checkIfYouCanMigrateData(wrk As Workbook) As Boolean
    checkIfYouCanMigrateData = False
    
    Dim sh As Worksheet
    Set sh = Nothing
    
    checkIfYouCanMigrateData = CBool( _
        checkOneSheet(wrk, SIXP.G_main_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_order_release_status_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_recent_build_plan_changes_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_cont_pnoc_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_osea_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_totals_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_resp_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_del_conf_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_open_issues_sh_nm, sh) _
        And checkOneSheet(wrk, SIXP.G_xq_sh_nm, sh) _
        )
    
End Function


Private Function checkOneSheet(wrk As Workbook, nameOfSheet As String, refSh As Worksheet) As Boolean
    
    checkOneSheet = False
    
    Set refSh = Nothing
    
    On Error Resume Next
    Set refSh = wrk.Sheets(CStr(nameOfSheet))
    
    If Not refSh Is Nothing Then
        ' also check first 4 columns
        
        If refSh.Cells(1, 1).Value = "Projekt" Then
            If refSh.Cells(1, 2).Value = "Plant Code" Then
                If refSh.Cells(1, 3).Value = "Faza" Then
                    If refSh.Cells(1, 4).Value = "CW" Then
                        checkOneSheet = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    
    checkOneSheet = False
    
End Function
    
    
Private Sub migrateDataBetween(sourceWorkbook As Workbook, destinationWorkbook As Workbook)


    ' lets the migration begin!
    ' -------------------------------------------------------------------
    
    ' main
    skopiujDane CStr(SIXP.G_main_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' ors
    skopiujDane CStr(SIXP.G_order_release_status_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' rbpc
    skopiujDane CStr(SIXP.G_recent_build_plan_changes_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' cont . pnoc
    skopiujDane CStr(SIXP.G_cont_pnoc_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' osea
    skopiujDane CStr(SIXP.G_osea_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' totals
    skopiujDane CStr(SIXP.G_totals_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' resp
    skopiujDane CStr(SIXP.G_resp_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' del conf
    skopiujDane CStr(SIXP.G_del_conf_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' open issues
    skopiujDane CStr(SIXP.G_open_issues_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    
    ' xq
    skopiujDane CStr(SIXP.G_xq_sh_nm), sourceWorkbook, destinationWorkbook
    SIXP.LoadingFormModule.incLoadingForm
    SIXP.LoadingFormModule.incLoadingForm
    
    
    ' -------------------------------------------------------------------
End Sub

Private Sub skopiujDane(shName As String, sWrk As Workbook, dWrk As Workbook)

    Dim sourceSh As Worksheet
    Dim destSh As Worksheet
    
    Set sourceSh = sWrk.Sheets(shName)
    Set destSh = dWrk.Sheets(shName)
    
    Dim destRange As Range
    Dim sourceRange As Range
    Dim lastRightSourceRange As Range
    
    
    
    ' logika przesuwajaca destRange do pierwszej pustej kolumny
    Set destRange = destSh.Range("A1")
    If Trim(destRange.Offset(1, 0)) <> "" Then
        Set destRange = destRange.End(xlDown).Offset(1, 0)
    Else
        Set destRange = destSh.Range("A2")
    End If
    
    
    ' source range zawsze od drugiej komorki lecimy po calosci
    Set sourceRange = sourceSh.Range("A2")
    Set lastRightSourceRange = sourceSh.Range("A1").End(xlToRight)
    
    
    ' proste sprawdzenie czy dane w ogole jakies sa w tym arkuszu
    ' nie zamierzam sie spuszczac nad problemami "Kocurowymi"
    ' ======================================================================
    If Trim(sourceRange.Value) <> "" Then
    
    
        ' logika kopiowania
        'Range("A7:N7").Select
        'Selection.Copy
        'Range("A10").Select
        'ActiveSheet.Paste
        Set sourceRange = zdefiniujKwadraciakaDoSkopiowaniaXD(sourceSh, sourceRange, lastRightSourceRange.Offset(1, 0))
        
        If Not sourceRange Is Nothing Then
        
            sourceRange.Copy
            destRange.PasteSpecial xlPasteAll
            
            'usun duplikaty teraz
            removeDuplicatesInDest destSh
            removeEmptyBetween destSh
            doSomeLeanFor destSh
            entireRowsRemovalIfEmpty destSh
        Else
            MsgBox "Kwadraciak nie zotal poprawnie zdefiniowany - proba migracji danych pomiedzy plikami 6P nie doszla do skutku! xD"
            End
        End If
        
    End If
    ' ======================================================================
    
    

End Sub

Private Sub removeEmptyBetween(dsh As Worksheet)

    Dim r As Range
    
    Set r = dsh.Cells(2 ^ 10, 1)
    Set r = r.End(xlUp)
    
    
    If r.Row > 1 Then
    
        Do
            If firstFourEmpty(r) Then
                r.EntireRow.Delete xlShiftUp
                
                
                Set r = dsh.Cells(2 ^ 10, 1)
                Set r = r.End(xlUp).Offset(1, 0)
            End If
        
            Set r = r.Offset(-1, 0)
        Loop Until r.Row = 1
    
    End If

End Sub

Private Sub doSomeLeanFor(dsh As Worksheet)



    If Trim(dsh.name) = SIXP.G_main_sh_nm Then
        ' do nothing for main - meh!
    Else

        Dim r As Range
        Set r = dsh.Cells(2, 1)
    
        Do
            If onlyFirstFourPotentialyFilled(r) Then
            
                ' 2 stepy: usuwania z arkusza danych plus usuniecie wpisu w arkuszu main
                If deleteThisFromMain(r) Then
            
                    r.EntireRow.Delete xlShiftUp
                    
                    Set r = dsh.Cells(1, 1)
                End If
            End If
            
            Set r = r.Offset(1, 0)
        Loop Until Trim(r) = ""
    End If

End Sub

Private Function deleteThisFromMain(r As Range) As Boolean
    deleteThisFromMain = False
    
    
    
    If Trim(r) <> "" Then
    
    
        Dim rm As Range, e As E_MAIN_ORDER
        Set rm = findProperRowInMainSheet(r)
        
        If Not rm Is Nothing Then
        
        
        
            If r.Parent.name = SIXP.G_order_release_status_sh_nm Then e = e_main_last_update_on_order_release_status
            If r.Parent.name = SIXP.G_recent_build_plan_changes_sh_nm Then e = e_main_last_update_on_recent_build_plan_changes
            If r.Parent.name = SIXP.G_cont_pnoc_sh_nm Then e = e_main_last_update_on_chart_contracted_pnoc
            If r.Parent.name = SIXP.G_osea_sh_nm Then e = e_main_last_update_on_osea
            If r.Parent.name = SIXP.G_totals_sh_nm Then e = e_main_last_update_on_totals
            If r.Parent.name = SIXP.G_xq_sh_nm Then e = e_main_last_update_on_xq
            If r.Parent.name = SIXP.G_del_conf_sh_nm Then e = e_main_last_update_on_del_conf
            If r.Parent.name = SIXP.G_open_issues_sh_nm Then e = e_main_last_update_on_open_issues
            If r.Parent.name = SIXP.G_resp_sh_nm Then e = e_main_last_update_on_resp
            
            
            rm.Offset(0, e - 1).Value = ""
            
            If rm.Offset(0, e - 1).Value = "" Then
                deleteThisFromMain = True
            Else
                deleteThisFromMain = False
            End If
            
        Else
        
            r.Interior.Color = RGB(250, 250, 50)
        
        End If
    End If
        
End Function

Private Function findProperRowInMainSheet(r As Range) As Range

    Set findProperRowInMainSheet = Nothing
    
    Dim l As T_Link
    Set l = New T_Link
    
    l.zrob_mnie_z_range r
    
    Set findProperRowInMainSheet = l.znajdz_siebie_w_arkuszu(ThisWorkbook.Sheets(SIXP.G_main_sh_nm))
End Function

Private Function firstFourEmpty(r As Range) As Boolean
    firstFourEmpty = False
    
    If Trim(r) = "" And Trim(r.Offset(0, 1)) = "" And Trim(r.Offset(0, 2)) = "" And Trim(r.Offset(0, 3)) = "" Then
        firstFourEmpty = True
    End If
End Function

Private Function onlyFirstFourPotentialyFilled(r As Range) As Boolean
    
    onlyFirstFourPotentialyFilled = False
    
    
    tmp = ""
    
    ' te 50 heurystycznie - tyle kolumn wydaje sie wystarczajace do sprawdzenia ilosci pustosci
    For x = 4 To 50
        tmp = tmp & Trim(r.Offset(0, x))
    Next x
    
    If Trim(tmp) = "" Then
        onlyFirstFourPotentialyFilled = True
    End If
    
End Function

Private Sub entireRowsRemovalIfEmpty(dsh As Worksheet)


    Dim r As Range
    Set r = dsh.Cells(1, 1)
    
    If r.Offset(1, 0) <> "" Then
        
        Set r = r.End(xlDown).Offset(1, 0)
        Set r = dsh.Range(r, dsh.Cells(2 ^ 10, 1))
        
        r.EntireRow.Delete xlShiftUp
    Else
    
        ' no operation
    End If
    
    
    

End Sub

Private Sub removeDuplicatesInDest(dsh As Worksheet)


    Application.DisplayAlerts = False


    'ActiveSheet.Range("$A$1:$N$549").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, _
        ' 7, 8, 9, 10, 11, 12, 13, 14), Header:=xlYes
        
    ' Array(1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39)
    dsh.Activate
    dsh.Cells(1, 1).Select
    dsh.Cells(1, 1).Activate
        
    Dim obszar As Range
    Set obszar = zdefiniujKwadraciakaDoSkopiowaniaXD(dsh, dsh.Cells(1, 1), dsh.Cells(1, 1).End(xlToRight))
    obszar.Select
    
    Dim lC As Integer
    lC = Int(dsh.Cells(1, 1).End(xlToRight).Column)
    
    ReDim varArr(lC - 1)
    Index = 0
    Do Until Index > Int(dsh.Cells(1, 1).End(xlToRight).Column) - 1
        varArr(Index) = Index + 1
        If Index + 1 = Int(dsh.Cells(1, 1).End(xlToRight).Column) Then Exit Do
        Index = Index + 1
    Loop
    
    obszar.RemoveDuplicates Columns:=(varArr), Header:=xlYes
    
    dsh.Activate
    dsh.Cells(1, 1).Select
    dsh.Cells(1, 1).Activate
    
    SIXP.LoadingFormModule.incLoadingForm
    
    
    Application.DisplayAlerts = True
End Sub

Private Function zdefiniujKwadraciakaDoSkopiowaniaXD(sh As Worksheet, topLeft As Range, topRight As Range) As Range

    Set zdefiniujKwadraciakaDoSkopiowaniaXD = Nothing
    
    
    
    Dim bottomRight As Range
    If topLeft.Offset(1, 0) <> "" Then
        Set bottomRight = sh.Cells(topLeft.End(xlDown).Row, topRight.Column)
    Else
        Set bottomRight = sh.Cells(topLeft.Row, topRight.Column)
    End If
    
    
    Set zdefiniujKwadraciakaDoSkopiowaniaXD = sh.Range(topLeft, bottomRight)
    
    SIXP.LoadingFormModule.incLoadingForm
    
End Function
    
