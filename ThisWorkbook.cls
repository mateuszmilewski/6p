VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    GLOBAL_POZWALAM_NA_REAKCJE_NA_ZMIANY = True
End Sub


Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)

    If G_ONE_ITEM_LOGIC_WAITING_FOR_SELECTION_CHANGE Then
    If FormGetOneItem.Visible Then
    
    
        ' MsgBox Target.Address
        FormGetOneItem.LabelAdr.Caption = Target.Address
        
        If CStr(FormGetOneItem.LabelClient.Caption) = CStr(FormOrderReleaseStatus.name) Then
        
        
            ' main or ors going back to global foo
            ' ----------------------------------------------------------------------------------
            
            tryToGetDataFromSelectionToForm Target, ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
            
            ' ----------------------------------------------------------------------------------
        ElseIf CStr(FormGetOneItem.LabelClient.Caption) = CStr(FormRecentBuildPlanChanges.name) Then
        
        
            ' main or rbpc going back to global foo
            ' ----------------------------------------------------------------------------------
            
            tryToGetDataFromSelectionToForm Target, ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
            
            ' ----------------------------------------------------------------------------------
        End If
        
    End If
    End If
End Sub
