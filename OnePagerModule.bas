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
    ' skonfiguruj_form_generowania_one_pagera ' to jest logika tylko dla starego formu
    
    FormOnePagerNEW.RadioPowerPoint.Value = True
    FormOnePagerNEW.Show vbModeless
    
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


Public Sub clear_new_one_pager()
    
'    'With ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM)
'
'    '
'    '
'    '    ' Q ORS
'    '    .Range("A5:H12").Value = ""
'    '
'    '    ' Q RBPC
'    '    .Range("J5:N11").Value = ""
'    '
'    '    ' fma and osea review
'    '    .Range("A16:F16").Value = ""
'    '
'    '    ' pus section
'    '    .Range("E17:E20").Value = ""
'    '
'    '    ' ppap section
'    '    .Range("E23:F23").Value = ""
'    '
'    '    ' ppap gate
'    '    .Range("G23").Value = "#"
'    '
'    '
'    '    ' osea section
'    '    .Range("A26:H26").Value = ""
'    '
'    '    ' resp
'    '    .Range("A29:H29").Value = ""
'    '
'    '
'    '    ' del conf
'    '    .Range("K16:M28").Value = ""
'    '
'    '
'    '
'    '
'    End With
    
    

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
        FormOnePager.ListBoxProjects.addItem k
    Next
    
    FormOnePager.ListBoxPlants.Clear
    For Each k In plt_dic.Keys
        FormOnePager.ListBoxPlants.addItem k
    Next
    
    FormOnePager.ListBoxPhases.Clear
    For Each k In faza_dic.Keys
        FormOnePager.ListBoxPhases.addItem k
    Next
    
    FormOnePager.ListBoxCWs.Clear
    
    
    putCwInOrder cw_dic, 0
    
    For Each k In cw_dic.Keys
        FormOnePager.ListBoxCWs.addItem k
    Next
    
    
    
    ' cfg na radio buttonach
    'FormOnePager.RadioPowerPoint.Value = False
    'FormOnePager.RadioExcels.Value = True
    
    With FormOnePager
        .RadioExcels.Value = False
        .RadioPowerPoint.Value = True
    End With
    
    FormOnePager.czy_uruchamiamy_eventy = True
    
    '
    ''
    ' -----------------------------------------------------------------------
    ' -----------------------------------------------------------------------
    
End Sub

Public Sub putCwInOrder(ByRef d As Dictionary, Optional startingValue As Long)

    ' ----------------------------------------------------------------------------------
    
    Dim x As Long, q As Long
    Dim refKey
    
    
    Dim mySortRange As Range
    Set mySortRange = ThisWorkbook.Sheets("register").Range("z1:aa1")
    
    mySortRange.EntireColumn.Delete
    Set mySortRange = ThisWorkbook.Sheets("register").Range("z1")
    Dim rRef As Range
    Set rRef = mySortRange.Offset(1, 0)
    For Each k In d.Keys
        
        rRef.Value = k
        rRef.Offset(0, 1).Value = startingValue
        Set rRef = rRef.Offset(1, 0)
    Next
    
    ' get entire range of yyyycw
    
    
    If ThisWorkbook.Sheets("register").Range("z1").Offset(2, 0) <> "" Then
        Set rRef = ThisWorkbook.Sheets("register").Range("z1").End(xlDown).End(xlDown)
    Else
        Set rRef = ThisWorkbook.Sheets("register").Range("z1").End(xlDown)
    End If
    
    
    Set mySortRange = ThisWorkbook.Sheets("register").Range(mySortRange, rRef.Offset(0, 1))
    mySortRange.Sort mySortRange, xlDescending
    
    
    Set d = Nothing
    Set d = New Dictionary
    Dim ir As Range
    For x = 1 To mySortRange.Count Step 2
    
        Set ir = mySortRange(x)
        ' Debug.Print ir.Row ' OK
        If ir <> "" Then
            d.Add ir, ir.Offset(0, 1)
        End If

            
    Next x
    
    
    ' ----------------------------------------------------------------------------------
End Sub
