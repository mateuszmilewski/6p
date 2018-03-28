Attribute VB_Name = "OverseaDataCatchModule"
Public Sub externalOverseaDataCatchLogic(valueFromList, frm As FormCatchWizard, labelText As String)
    innerSub valueFromList, frm, labelText
End Sub





Private Sub innerSub(valueFromList, frm As FormCatchWizard, labelText As String)

    
    
    
    
    
    FormOverseaDataCatcher.activeSourceSheetName = CStr(valueFromList)
    FormOverseaDataCatcher.labelka = CStr(labelText)
    FormOverseaDataCatcher.BtnSubmit.Caption = "GET DATA"
    FormOverseaDataCatcher.Show vbModeless
    Workbooks(CStr(valueFromList)).Activate
    
End Sub

