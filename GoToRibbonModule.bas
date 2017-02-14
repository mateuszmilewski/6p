Attribute VB_Name = "GoToRibbonModule"
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

Public Sub goto_wiz_buff_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Activate
End Sub

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

Public Sub goto_xq_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_xq_sh_nm).Activate
End Sub

Public Sub goto_one_pager_sh(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm).Activate
End Sub
