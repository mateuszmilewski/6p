VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewOnePagerHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private kolekcja_linkow As Collection
Private link_raportu As T_Link
Private my_linker As Linker


Private m As Worksheet
Private pcs(1 To 9) As Worksheet

Private one_pager_sheet As Worksheet


Private ppApp As PowerPoint.Application
Private pp As Presentation
Private ppSlide As PowerPoint.Slide
Private myCustomLayout As PowerPoint.CustomLayout



Public Function get_one_pager_sheet() As Worksheet
    Set get_one_pager_sheet = one_pager_sheet
End Function


Private Sub Class_Initialize()
    
    Set kolekcja_linkow = New Collection
    
    
    
    ' heading
    ' 1 == main enum
    Set pcs(1) = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' left side
    ' 2 == order release status enum
    Set pcs(2) = ThisWorkbook.Sheets(SIXP.G_order_release_status_sh_nm)
    Set pcs(3) = ThisWorkbook.Sheets(SIXP.G_cont_pnoc_sh_nm)
    Set pcs(4) = ThisWorkbook.Sheets(SIXP.G_osea_sh_nm)
    
    ' right side
    ' 5 == recent build plan changes enum
    Set pcs(5) = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
    Set pcs(6) = ThisWorkbook.Sheets(SIXP.G_totals_sh_nm)
    Set pcs(7) = ThisWorkbook.Sheets(SIXP.G_del_conf_sh_nm)
    
    ' bottom
    ' 8 == resp enum
    Set pcs(8) = ThisWorkbook.Sheets(SIXP.G_resp_sh_nm)
    
    ' second page
    ' 9 == open issues enum
    Set pcs(9) = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
    
    
    Set one_pager_sheet = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
    
    Set chart1 = ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm)
    Set chart2 = ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm)
    Set chart3 = ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm)
End Sub

Private Sub Class_Terminate()
    Set kolekcja_linkow = Nothing
    
    
    For x = 1 To 9
        Set pcs(x) = Nothing
    Next x
End Sub
