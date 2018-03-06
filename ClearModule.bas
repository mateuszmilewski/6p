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
    
    
    If SIXP.GlobalFooModule.checkIfFirstFourFieldsProjektPlantCodeFazaCW(ActiveSheet) Then
        
        If CStr(ActiveSheet.name) = CStr(SIXP.G_main_sh_nm) Then
        
            
        
            If ActiveCell.Column > 1 And ActiveCell.Row > 4 And ActiveCell.Value <> "" Then
            
                MsgBox "not imeplemented yet"
            Else
                MsgBox "nie mozesz dla tej selekcji nic usunac!"
            End If
        Else
        
            If ActiveCell.Column > 1 And ActiveCell.Value <> "" Then
            
                MsgBox "not imeplemented yet"
            Else
                MsgBox "nie mozesz dla tej selekcji nic usunac!"
            End If
        
        End If
    End If
End Sub
    
