Attribute VB_Name = "GoToRibbonModule"
Public Sub goto_register_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Activate
End Sub

Public Sub goto_ors_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm).Activate
End Sub

Public Sub goto_cp_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm).Activate
End Sub

Public Sub goto_osea_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_osea_sh_nm).Activate
End Sub

Public Sub goto_rbpc_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm).Activate
End Sub

Public Sub goto_main_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_main_sh_nm).Activate
End Sub

Public Sub goto_resp_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_resp_sh_nm).Activate
End Sub

Public Sub goto_oi_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Activate
End Sub

Public Sub goto_cfg_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_config_sh_nm).Activate
End Sub

Public Sub goto_tot_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_totals_sh_nm).Activate
End Sub

Public Sub goto_del_conf_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm).Activate
End Sub
