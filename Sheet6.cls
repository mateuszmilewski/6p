VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    If Target.Column > 5 Then
        ' jestesmy na polach odpowiadajacych za linkowanie z informacjami dostepnymi w arkuszyach side'owych
        ' ===================================================================================================
        ' ===================================================================================================
        ''
        '
        
        Dim tl_main As T_Link, tl As T_Link
        Set tl_main = New T_Link
        tl_main.zrob_mnie_z_range Target.Parent.Cells(Target.Row, 1)
        
        
        If Target.Column = SIXP.e_main_last_update_on_order_release_status Then
            ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_recent_build_plan_changes Then
            ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_chart_contracted_pnoc Then
            ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_osea Then
            ThisWorkbook.Sheets(SIXP.G_osea_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_totals Then
            ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_xq Then
            ThisWorkbook.Sheets(SIXP.G_xq_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_del_conf Then
            ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_open_issues Then
            ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Activate
            
        ElseIf Target.Column = SIXP.e_main_last_update_on_resp Then
            ThisWorkbook.Sheets(SIXP.G_resp_sh_nm).Activate
            
        End If
        
        If Not tl_main.znajdz_siebie_w_arkuszu(ActiveSheet) Is Nothing Then
            tl_main.znajdz_siebie_w_arkuszu(ActiveSheet).Select
        Else
            MsgBox "rekord nie istnieje! Wracam do arkusza MAIN"
            ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
        End If
        
        '
        ''
        ' ===================================================================================================
        ' ===================================================================================================
    End If
End Sub
