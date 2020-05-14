Attribute VB_Name = "ImportFromWizardModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
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

Public Sub import_wizard_content(ictrl As IRibbonControl)

    ' usuniecie danych z wizard buff
    ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM).Range("a1:zz1000").Clear
    
    FormCatchWizard.ListBox1.Clear
    FormCatchWizard.ListBox1.MultiSelect = fmMultiSelectSingle
    
    With FormCatchWizard
        .BtnImportOpenIssues.Enabled = False
        .BtnJustImport.Enabled = True
        .BtnSubmit.Enabled = True
        .BtnGetFrom6P.Enabled = False
        .BtnOsea.Enabled = False
    End With
        
    For Each w In Workbooks
        With FormCatchWizard.ListBox1
            .addItem w.name
        End With
    Next w
    
    FormCatchWizard.Show vbModeless

End Sub
