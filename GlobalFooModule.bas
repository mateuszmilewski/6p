Attribute VB_Name = "GlobalFooModule"
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


    If CStr(Trim(yyyycw)) = "" Then
        from_yyyy_cw_to_monday_from_this_week = Date
    Else
    
        Y = Left(yyyycw, 4)
        cw = Right(yyyycw, 2)
            
        ' -------------------- ' -------------------- ' --------------------
        
        Dim d As Date
        d = CDate(Y & "-01-01")
        
        Do
            d = d + 1
        Loop Until CLng(Application.WorksheetFunction.IsoWeekNum(CDbl(d))) = CLng(cw)
        
        from_yyyy_cw_to_monday_from_this_week = d
    End If
End Function

Public Sub global_goto_main_sheet_with_selection_on_data(r As Range)

    If r.Column <= 4 Then
        ' niezaleznie od arkusza zawsze pierwsze 4 kolumny to link
        
        Dim tl As T_Link
        Set tl = New T_Link
        
        tl.zrob_mnie_z_range r.Parent.Cells(r.Row, 1)
        
        If Not tl.znajdz_siebie_w_arkuszu(ThisWorkbook.Sheets(SIXP.G_main_sh_nm)) Is Nothing Then
            ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
            tl.znajdz_siebie_w_arkuszu(ActiveSheet).Select
        Else
            MsgBox "rekord nie istnieje!"
            
        End If
    End If
End Sub

Public Sub global_form_openers(r As Range)
    
    If r.Column > 4 Then
        If Trim(r.Parent.Cells(r.Row, SIXP.e_link_project)) <> "" Then
        
            ' ----------------------------------------------------------
            ''
            '
            
            ' mamy tutaj dwa obiekty: jeden typu T_Link, drugi typu Linker
            ' tak jak czesto gesto korzytam z samego Linku, tak linker pozostawal
            ' w tyle, jednak teraz w tej logice wrocil do lask
            ' jeszcze kwestia filozoficznam czy przypadkiem Linker nie powinien ze swoimi
            ' metodami nie byc czescia jako komponent obiektu typu T_Link
            ' ale powiem szczerze nie chce mi sie juz tego reimplemntowac ponownie
            Dim used_to_be_txt_from_combobox As String
            
            Dim l As T_Link
            Set l = New T_Link
            Dim lr As Linker
            Set lr = New Linker
            ' l.zrob_mnie_z_argsow proj_txt, Me.TextBoxPlt, Me.TextBoxFaza, Me.TextBoxCW
            l.zrob_mnie_z_argsow CStr(r.Parent.Cells(r.Row, SIXP.e_link_project)), _
                CStr(r.Parent.Cells(r.Row, SIXP.e_link_plt)), _
                CStr(r.Parent.Cells(r.Row, SIXP.e_link_faza)), _
                CStr(r.Parent.Cells(r.Row, SIXP.e_link_cw))
                
            used_to_be_txt_from_combobox = CStr(lr.return_full_concated_r_string_comma_seperated(l))
            
            If r.Parent.name = SIXP.G_order_release_status_sh_nm Then
                
                SIXP.DataFlowPodFormMainModule.zrob_order_release_status _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_cont_pnoc_sh_nm Then
                
                SIXP.DataFlowPodFormMainModule.zrob_contracted_pnoc _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_del_conf_sh_nm Then
                
                SIXP.DataFlowPodFormMainModule.zrob_del_conf _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_open_issues_sh_nm Then
                
                SIXP.DataFlowPodFormMainModule.zrob_open_issues _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_osea_sh_nm Then
                
                SIXP.DataFlowPodFormMainModule.zrob_osea_scope _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_recent_build_plan_changes_sh_nm Then
                
                SIXP.DataFlowPodFormMainModule.zrob_recent_build_plan_changes _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_resp_sh_nm Then
            
                SIXP.DataFlowPodFormMainModule.zrob_resp _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_totals_sh_nm Then
            
                SIXP.DataFlowPodFormMainModule.zrob_total _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                    
            ElseIf r.Parent.name = SIXP.G_xq_sh_nm Then
            
                SIXP.DataFlowPodFormMainModule.zrob_xq _
                    SIXP.G_BTN_TEXT_EDIT, CStr(used_to_be_txt_from_combobox)
                
            End If
            
            '
            ''
            ' ----------------------------------------------------------
        End If
    End If
End Sub

Public Function global_check_if_empty_and_put_zero(s As String) As String
    
    If Trim(s) = "" Then
        global_check_if_empty_and_put_zero = "0"
    Else
        global_check_if_empty_and_put_zero = Trim(s)
    End If
End Function

' alias
Public Function global_cpz(s As String) As String
    global_cpz = CStr(global_check_if_empty_and_put_zero(CStr(s)))
End Function



Public Sub gotoThisWorkbookMainA1()
    
    ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
    ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Cells(1, 1).Select
End Sub


Public Function checkIfFirstFourFieldsProjektPlantCodeFazaCW(ByRef sh As Worksheet) As Boolean

    checkIfFirstFourFieldsProjektPlantCodeFazaCW = False
    
    If sh.Cells(1, 1).Value = "Projekt" Then
        If sh.Cells(1, 2).Value = "Plant Code" Then
            If sh.Cells(1, 3).Value = "Faza" Then
                If sh.Cells(1, 4).Value = "CW" Then
                        
                        checkIfFirstFourFieldsProjektPlantCodeFazaCW = True
                            
                End If
            End If
        End If
    End If
End Function





Public Sub copyOneItemFromDifferentRecord(frmName As String)
    
    Dim r As Range
    Set r = ActiveCell
    
    ' MsgBox r.Address
    G_ONE_ITEM_LOGIC_WAITING_FOR_SELECTION_CHANGE = True
    
    FormGetOneItem.LabelClient.Caption = CStr(frmName)
    FormGetOneItem.Show vbModeless
End Sub


Public Sub tryToGetDataFromSelectionToForm(r As Range, sh As Worksheet)


    ' r - selection
    ' sh - worksheet with data

    ' a few cells in this range connected with searched dataset
    Dim rr As Range
    
    Dim l As T_Link
    Set l = New T_Link
    l.zrob_mnie_z_range (r.Parent.Cells(r.Row, 1))
    Set rr = l.znajdz_siebie_w_arkuszu(sh)

    FormGetOneItem.LabelData.Caption = CStr(fillOneItemFromSelection(rr))

End Sub

Private Function fillOneItemFromSelection(r As Range) As String


    ' arg r is on first column in dedicated worksheet

    fillOneItemFromSelection = ""
    
    If r.Parent.name = SIXP.G_order_release_status_sh_nm Then
    
        With FormGetOneItem
            .ors_mrd = r.Offset(0, SIXP.e_order_release_mrd - 1).Value
            .ors_bom_freeze = r.Offset(0, SIXP.e_order_release_bom_freeze - 1).Value
            .ors_build = r.Offset(0, SIXP.e_order_release_build - 1).Value
            .ors_no_of_veh = r.Offset(0, SIXP.e_order_release_no_of_veh - 1).Value
            .ors_orders_due = r.Offset(0, SIXP.e_order_release_orders_due - 1).Value
            .ors_released = r.Offset(0, SIXP.e_order_release_released - 1).Value
            .ors_weeks_delay = r.Offset(0, SIXP.e_order_release_weeks_delay - 1).Value
            
            
            .clear_rbpc
            
            
            fillOneItemFromSelection = "data from " & SIXP.G_order_release_status_sh_nm & " is now stored!"
        End With
    
    ElseIf r.Parent.name = SIXP.G_recent_build_plan_changes_sh_nm Then
    
    
        With FormGetOneItem
        
        
            .rbpc_comment = r.Offset(0, SIXP.e_recent_bp_ch_comment - 1).Value
            .rbpc_num_of_veh = r.Offset(0, SIXP.e_recent_bp_ch_no_of_veh - 1).Value
            .rbpc_order_release_changes = r.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1).Value
            .rbpc_tbw = r.Offset(0, SIXP.e_recent_bp_ch_tbw - 1).Value
            
            
            .clear_ors
            
            fillOneItemFromSelection = "data from " & SIXP.G_recent_build_plan_changes_sh_nm & " is now stored!"
        End With
    End If
    

    
        
End Function
