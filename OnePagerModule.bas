Attribute VB_Name = "OnePagerModule"
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

Public Sub generate_one_pager(ictrl As IRibbonControl)
    clear_one_pager
    skonfiguruj_form_generowania_one_pagera
    FormOnePager.Show
    
    MsgBox "ready"
End Sub


Public Sub clear_one_pager()
    
    With ThisWorkbook.Sheets(SIXP.G_one_pager_sh_nm)
    
        ' 1st i 2nd pcs
        .Range("A3:Z10") = ""
        
        'PEM
        .Range("C37") = ""
        
        ' SQE
        .Range("C39") = ""
        
        ' PPM
        .Range("M37") = ""
        
        ' FMA
        .Range("W37") = ""
        
        ' OSEA
        .Range("M39") = ""
        
        ' open issues
        .Range("AG3:AP25") = ""
        
        
        ' zera na del confie
        ' ------------------
        
        ' calci after & before
        .Range("AA25:AA31") = 0
        .Range("AC25:AC31") = 0
        
        
        ' green
        .Range("T25:T28") = 0
        
        
        ' red
        .Range("T30:T32") = 0
        
        
        'sumy
        ' sa fromuly na stale
        '.Range("U34") = 0
        '.Range("X34") = 0
        '.Range("AA34") = 0
        ' ------------------
        
    End With
    
    
    
    ' chart 1 == pnoc
    With ThisWorkbook.Sheets(SIXP.G_chart1_handler_sh_nm)
        ' buff na chartach
        ' ------------------
        
        
        ' dokladnie tyle ile trzeba
        ' .Range("D6:F12").Clear
        ' dam wiecej tak just in case
        .Range("D6:F100").Clear
        
        ' ------------------
    End With
    
    
    
    
    ' chart 2 == OSEA
    With ThisWorkbook.Sheets(SIXP.G_chart2_handler_sh_nm)
        ' buff na chartach
        ' ------------------
        
        .Range("C6:I6") = 0
        
        ' ------------------
    End With
    
    
    
    ' chart 3 == total
    With ThisWorkbook.Sheets(SIXP.G_chart3_handler_sh_nm)
        ' buff na chartach
        ' ------------------
        
        .Range("C6:F6") = 0
        
        .Range("G6:G7") = 0
        
        .Range("I6") = 0
        
        ' arrived ' in transit ' future
        .Range("K6:K8") = 0
        
        
        ' ppap / no ppap
        .Range("M6:N6") = 0
        
        ' ------------------
    End With
End Sub

Public Sub skonfiguruj_form_generowania_one_pagera()
    
    
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
    ''
    '
    Dim proj_dic As Dictionary
    Dim plt_dic As Dictionary
    Dim faza_dic As Dictionary
    Dim cw_dic As Dictionary
    
    Dim m As Worksheet
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    Dim r As Range
    Set r = m.Cells(2, 1)
    
    Set proj_dic = New Dictionary
    Set plt_dic = New Dictionary
    Set faza_dic = New Dictionary
    Set cw_dic = New Dictionary
    
    'proj_dic.Add r, 1
    'plt_dic.Add r.Offset(0, 1), 1
    'faza_dic.Add r.Offset(0, 2), 1
    'cw_dic.Add r.Offset(0, 3), 1
    
    Do
        If Not proj_dic.Exists(Trim(r)) Then
            proj_dic.Add Trim(r), 1
        End If
        
        If Not plt_dic.Exists(Trim(r.Offset(0, 1))) Then
            plt_dic.Add Trim(r.Offset(0, 1)), 1
        End If
        
        If Not faza_dic.Exists(Trim(r.Offset(0, 2))) Then
            faza_dic.Add Trim(r.Offset(0, 2)), 1
        End If
        
        If Not cw_dic.Exists(Trim(r.Offset(0, 3))) Then
            cw_dic.Add Trim(r.Offset(0, 3)), 1
        End If
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    FormOnePager.ListBoxProjects.Clear
    For Each k In proj_dic.Keys
        FormOnePager.ListBoxProjects.AddItem k
    Next
    
    FormOnePager.ListBoxPlants.Clear
    For Each k In plt_dic.Keys
        FormOnePager.ListBoxPlants.AddItem k
    Next
    
    FormOnePager.ListBoxPhases.Clear
    For Each k In faza_dic.Keys
        FormOnePager.ListBoxPhases.AddItem k
    Next
    
    FormOnePager.ListBoxCWs.Clear
    For Each k In cw_dic.Keys
        FormOnePager.ListBoxCWs.AddItem k
    Next
    
    
    
    ' cfg na radio buttonach
    FormOnePager.RadioPowerPoint.Value = False
    FormOnePager.RadioExcels.Value = True
    
    FormOnePager.czy_uruchamiamy_eventy = True
    
    '
    ''
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
    
End Sub
