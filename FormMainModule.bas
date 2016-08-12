Attribute VB_Name = "FormMainModule"
Public Sub run_FormMain(Optional link_str As String)


    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 1

    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.init link_str
    
    Set fmh = Nothing
    
    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 0
End Sub

Public Sub add_new_project()

    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 1
    
    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.new_project
    
    Set fmh = Nothing
    
    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 0
End Sub







Public Sub adjustuj_guzik(e_main As E_MAIN_ORDER, ish As Worksheet, imainsh As Worksheet, il As T_Link, paleta As PaletaTheDailyCommute)



    ' dzieki koncepcji "porownania posredniego" mamy latwosc z ogarnieciem
    ' jaki kolor dany button powinien miec
    ' wydzielilem podwojnie logike a w sumie cale dzialo te funkcji mozna byloby rozdzielic
    ' od razu na wysokosci suba ktory sie znajduje powyzej
    ' byloby nieco mniej kodu, ale skoro juz to napisalem
    ' to niech to tak zostanie - szkoda czasu i nie bede sobie marnowal statsow
    
    
    Dim range_from_main As Range, range_from_ish As Range
    
    ' porownanie posrednie
    Set range_from_main = il.znajdz_siebie_w_arkuszu(imainsh)
    Set range_from_ish = il.znajdz_siebie_w_arkuszu(ish)
    
    ' jesli te dwie zmienne nie sa puste to znaczy ze mamy takie dane spasowane i chcemy je edytowac
    If Not range_from_main Is Nothing Then
        If Not range_from_ish Is Nothing Then
            
            If e_main = e_main_last_update_on_order_release_status Then
            
                With FormMain.BtnOrderReleaseStatus
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_recent_build_plan_changes Then
                
                With FormMain.BtnRecentBuildPlanChanges
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_chart_contracted_pnoc Then
                
                With FormMain.BtnContractedPNOC
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_osea Then
                
                With FormMain.BtnOseaScope
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_totals Then
                
                With FormMain.BtnTotals
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_xq Then
                
                With FormMain.BtnXq
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_del_conf Then
                
                With FormMain.BtnDelConf
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_open_issues Then
                
                With FormMain.BtnOpenIssues
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_resp Then
                
                With FormMain.BtnResp
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            End If
        Else
            ' sekcja, gdzie cos znalezlismy w main jednak nie ma tego w arkuszu przeszukiwanym
            ' to znaczy tyle ze trzeba dodac nowe
            
            If e_main = e_main_last_update_on_order_release_status Then
            
                FormMain.BtnOrderReleaseStatus.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOrderReleaseStatus.BackColor = paleta.yellow
                FormMain.BtnOrderReleaseStatus.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_recent_build_plan_changes Then
            
                FormMain.BtnRecentBuildPlanChanges.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnRecentBuildPlanChanges.BackColor = paleta.yellow
                FormMain.BtnRecentBuildPlanChanges.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_chart_contracted_pnoc Then
            
                FormMain.BtnContractedPNOC.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnContractedPNOC.BackColor = paleta.yellow
                FormMain.BtnContractedPNOC.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_osea Then
            
                FormMain.BtnOseaScope.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOseaScope.BackColor = paleta.yellow
                FormMain.BtnOseaScope.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_totals Then
            
                FormMain.BtnTotals.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnTotals.BackColor = paleta.yellow
                FormMain.BtnTotals.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_xq Then
            
                FormMain.BtnXq.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnXq.BackColor = paleta.yellow
                FormMain.BtnXq.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_del_conf Then
            
                FormMain.BtnDelConf.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnDelConf.BackColor = paleta.yellow
                FormMain.BtnDelConf.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_open_issues Then
            
                FormMain.BtnOpenIssues.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOpenIssues.BackColor = paleta.yellow
                FormMain.BtnOpenIssues.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_resp Then
            
                FormMain.BtnResp.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnResp.BackColor = paleta.yellow
                FormMain.BtnResp.ForeColor = paleta.dark_grey
                
            End If
        End If
    End If
    
    
    
    
End Sub


