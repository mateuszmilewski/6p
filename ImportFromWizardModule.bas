Attribute VB_Name = "ImportFromWizardModule"
Public Sub import_wizard_content(ictrl As IRibbonControl)

    ' usuniecie danych z wizard buff
    ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("a1:zz1000").Clear
    
    FormCatchWizard.ListBox1.Clear
    
    For Each w In Workbooks
        With FormCatchWizard.ListBox1
            .AddItem w.Name
        End With
    Next w
    
    FormCatchWizard.Show

End Sub
