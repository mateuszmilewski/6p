VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormMain 
   Caption         =   "Main Interface"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   OleObjectBlob   =   "FormMain.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnContractedPNOC_Click()
    Hide
    zrob_contracted_pnoc CStr(Me.BtnContractedPNOC.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnDelConf_Click()
    Hide
    SIXP.DataFlowPodFormMainModule.zrob_del_conf CStr(Me.BtnDelConf.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnOpenIssues_Click()
    Hide
    SIXP.DataFlowPodFormMainModule.zrob_open_issues CStr(Me.BtnOpenIssues.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnOrderReleaseStatus_Click()
    Hide
    zrob_order_release_status CStr(Me.BtnOrderReleaseStatus.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnOseaScope_Click()
    Hide
    zrob_osea_scope CStr(Me.BtnOseaScope.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnRecentBuildPlanChanges_Click()
    Hide
    zrob_recent_build_plan_changes CStr(Me.BtnRecentBuildPlanChanges.Caption), CStr(Me.ComboBoxProject.Value)
End Sub




Private Sub BtnResp_Click()
    Hide
    SIXP.zrob_resp CStr(Me.BtnResp.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnTotals_Click()
    Hide
    SIXP.zrob_total CStr(Me.BtnTotals.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub BtnXq_Click()
    Hide
    SIXP.zrob_xq CStr(Me.BtnXq.Caption), CStr(Me.ComboBoxProject.Value)
End Sub

Private Sub ComboBoxProject_Change()




    If Me.Visible = True Then

        Dim l As T_Link
        Dim sh As Worksheet
        Dim main_sh As Worksheet
        Set main_sh = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
        
        Dim paleta As PaletaTheDailyCommute
        Set paleta = New PaletaTheDailyCommute
        
        
        arr = Split(Me.ComboBoxProject.Value, ",")
        Set l = New T_Link
        l.zrob_mnie_z_argsow Trim(arr(0)), Trim(arr(1)), Trim(arr(2)), Trim(arr(3))
        ' wartosc w combo box okreslone
        ' czas na texty w guzikach
        
        ' order release status
        Set sh = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_order_release_status, sh, main_sh, l, paleta
        
        ' recent bp ch
        Set sh = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_recent_build_plan_changes, sh, main_sh, l, paleta
        
        ' chart cont pnoc
        Set sh = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_chart_contracted_pnoc, sh, main_sh, l, paleta
        
        ' sea
        Set sh = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_osea, sh, main_sh, l, paleta
        
        ' totals
        Set sh = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_totals, sh, main_sh, l, paleta
        
        ' xq
        Set sh = ThisWorkbook.Sheets(SIXP.G_xq_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_xq, sh, main_sh, l, paleta
        
        ' del conf
        Set sh = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_del_conf, sh, main_sh, l, paleta
        
        ' open issues
        Set sh = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_open_issues, sh, main_sh, l, paleta
        
        ' resp
        Set sh = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
        adjustuj_guzik SIXP.e_main_last_update_on_resp, sh, main_sh, l, paleta
    End If
End Sub
