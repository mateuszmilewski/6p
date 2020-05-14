Attribute VB_Name = "SechelModule"
Public Sub parse_sechel_export(ictrl As IRibbonControl)


    With FormGetExtractFromSechel
        
        .ListBox1.Clear
        .DTPicker1.Value = CDate(Date)
        
        Dim w As Workbook
        For Each w In Application.Workbooks
        
            .ListBox1.addItem w.name
        Next w

        .Show
    End With
End Sub



Public Sub fill_details_from_parsed_data(ictrl As IRibbonControl)


    Dim r As Range
    Dim l As T_Link
    Set l = New T_Link
    
    
    
    Dim referencjaMain As Range
    Dim mainSh As Worksheet
    
    If Selection.Count = 1 Then
    
        
        l.zrob_mnie_z_range Selection.Parent.Cells(Selection.Row, 1)
        Set referencjaMain = l.znajdz_siebie_w_arkuszu(ThisWorkbook.Sheets(SIXP.G_main_sh_nm))
        
        
        If referencjaMain Is Nothing Then
            MsgBox "Brak danych referencyjnych!"
            Exit Sub
        Else
        
            referencjaMain.Parent.Activate
            Set mainSh = referencjaMain.Parent
            ' mainSh.ShowAllData
            ' referencjaMain.Select
            
            
            With FormSechel
            
                
                ' clear schowek
                .LabelSchowek.Caption = ""
            
                
                ' main link for project
                ' ---------------------------------------------------------------
                .Label_REF_ADR.Caption = referencjaMain.Address
                .Label_REF_PROJ.Caption = referencjaMain.Value
                .Label_REF_PLT.Caption = referencjaMain.Offset(0, 1).Value
                .Label_REF_FAZA.Caption = referencjaMain.Offset(0, 2).Value
                .Label_REF_YYYYCW.Caption = referencjaMain.Offset(0, 3).Value
                ' ---------------------------------------------------------------
                
                
                ' sechel buffer
                ' ---------------------------------------------------------------
                
                .TextBoxLines.Value = findInSechelBuffCell("LINES")
                
                .TextBoxRecu.Value = findInSechelBuffCell("RECU")
                .TextBoxFauxManquant.Value = findInSechelBuffCell("FauxManquant")
                .TextBoxManquantPlus.Value = findInSechelBuffCell("manquantPlus")
                .TextBoxAVenir.Value = findInSechelBuffCell("A venir")
                .TextBoxEnCours.Value = findInSechelBuffCell("en cours")
                .TextBoxManquant.Value = findInSechelBuffCell("manquant")
                
                
                ' ---------------------------------------------------------------
                
                
                
                ' ORDER RELEASE STATUS
                ' ---------------------------------------------------------------
                
                Dim rRef As Range
                Set rRef = l.znajdz_siebie_w_arkuszu(ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm))
                
                
                If Not rRef Is Nothing Then
                    
                    Dim enum_ors As E_ORDER_RELEASE_STATUS
                    
                    enum_ors = e_order_release_mrd
                    .TextBox_ORS_MRD.Value = rRef.Offset(0, enum_ors - 1).Value
                    
                    enum_ors = e_order_release_build
                    .TextBox_ORS_Build.Value = rRef.Offset(0, enum_ors - 1).Value
                    
                    enum_ors = e_order_release_bom_freeze
                    .TextBox_ORS_BOMfreeze.Value = rRef.Offset(0, enum_ors - 1).Value
                    
                    enum_ors = e_order_release_no_of_veh
                    .TextBox_ORS_noOfVeh.Value = rRef.Offset(0, enum_ors - 1).Value
                    
                    enum_ors = e_order_release_orders_due
                    .TextBox_ORS_OrdersDue.Value = rRef.Offset(0, enum_ors - 1).Value
                    
                    enum_ors = e_order_release_released
                    .TextBox_ORS_Released.Value = rRef.Offset(0, enum_ors - 1).Value
                    
                    enum_ors = e_order_release_weeks_delay
                    .TextBox_ORS_weeksDelay.Value = rRef.Offset(0, enum_ors - 1).Value
                
                
                End If
                
                
                
                
                Set rRef = l.znajdz_siebie_w_arkuszu(ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm))
                
                
                If Not rRef Is Nothing Then
                
                    Dim e_rbpc As E_RECENT_BP_CH
                    
                    e_rbpc = e_recent_bp_ch_no_of_veh
                    .TextBox_RBPC_numOfVeh.Value = rRef.Offset(0, e_rbpc - 1).Value
                    
                    e_rbpc = e_recent_bp_ch_tbw
                    .TextBox_RBPC_TBW.Value = rRef.Offset(0, e_rbpc - 1).Value
                    
                    e_rbpc = e_recent_bp_ch_order_release_ch
                    .TextBox_RBPC_orderReleaseChanges.Value = rRef.Offset(0, e_rbpc - 1).Value
                    
                    e_rbpc = e_recent_bp_ch_comment
                    .TextBox_RBPC_Comment.Value = rRef.Offset(0, e_rbpc - 1).Value
                End If
                
                
                
                Set rRef = l.znajdz_siebie_w_arkuszu(ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm))
                
                If Not rRef Is Nothing Then
                
                    Dim e_cont_pnoc As E_CONT_PNOC_CHART
                    
                    e_cont_pnoc = e_cont_pnoc_chart_contracted
                    .TextBox_3_Contracted.Value = rRef.Offset(0, e_cont_pnoc - 1).Value
                    
                    e_cont_pnoc = e_cont_pnoc_chart_pnoc
                    .TextBox_3_PNOC.Value = rRef.Offset(0, e_cont_pnoc - 1).Value
                    
                    e_cont_pnoc = e_cont_pnoc_chart_open_bp
                    .TextBox_3_OpenBP.Value = rRef.Offset(0, e_cont_pnoc - 1).Value
                    
                    e_cont_pnoc = e_cont_pnoc_chart_actionable_fma
                    .TextBox_3_actionableFMA.Value = rRef.Offset(0, e_cont_pnoc - 1).Value
                End If
                
                
                ' ---------------------------------------------------------------
                
                
                .Show vbModelss
            End With
        
        
        End If
        
    End If
End Sub



Private Function findInSechelBuffCell(ptrn As String)


    findInSechelBuffCell = ""

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(SIXP.G_SECHEL_BUFF_SH_NM)
    
    Dim r As Range
    Set r = sh.Range("A1")
    
    Do
        If Trim(r.Value) = CStr(ptrn) Then
            findInSechelBuffCell = CStr(r.Offset(0, 1).Value)
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function
