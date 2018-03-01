VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeletionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2017 Mateusz Forrest Milewski
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


Public Sub usun_wszystko()
    
    stale_wygodne_odniesienie = "A2:AA500000"
    strrows = "2:500000"
    
    
    
    Dim sh As Worksheet, rng As Range
    For Each sh In ThisWorkbook.Sheets
    
        On Error Resume Next
        sh.ShowAllData
        
    
        Set rng = sh.Cells(1, 1)
        
        If rng.Value = "Projekt" Then
            If rng.Offset(0, 1) = "Plant Code" Then
                If rng.Offset(0, 2) = "Faza" Then
                    If rng.Offset(0, 3) = "CW" Then
                    
                        
                        
                    
                        sh.Range(stale_wygodne_odniesienie).ClearComments
                        sh.Range(stale_wygodne_odniesienie).Value = ""
                        sh.Rows(strrows).Delete xlUp
                    End If
                End If
            End If
        End If
    Next sh
End Sub


Public Sub usun_po_gwiazdce(pattern As String)


    If Len(CStr(Trim(pattern))) > 0 Then
    
        ans = MsgBox("Czy chces usunac kazde wystapienie itemu zawierajace: " & _
            Trim(CStr(pattern)), vbYesNo)
            
        If ans = vbYes Then
        
            strlink = pattern
        
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_main_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_cont_pnoc_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_del_conf_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_open_issues_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_order_release_status_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_osea_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_recent_build_plan_changes_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_resp_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_totals_sh_nm), CStr(strlink), True
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_xq_sh_nm), CStr(strlink), True
        Else
            MsgBox "makro konczy prace"
            End
        End If
    
    Else
    
        MsgBox "jakims cudem ten msgbox pojawil sie - zawolaj milewskiego!"
        End
    End If
    
End Sub

Public Sub usun_kazde_wystapienie_dla_aktywnej_komorki(r As Range)


    If r.Row >= 2 And Trim(r.Parent.Cells(r.Row, 1)) <> "" Then

        ans = MsgBox("Czy chces usunac kazde wystapienie itemu: " & _
            Trim(CStr(r.Parent.Cells(r.Row, 1))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 2))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 3))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 4))), vbYesNo)
            
        If ans = vbYes Then
            ' BODY USUWANIA DANYCH Z WSZYTKICH ARKUSZY
            ' =========================================================
            ' =========================================================
            ''
            '
            strlink = Trim(CStr(r.Parent.Cells(r.Row, 1))) & _
                ", " & Trim(CStr(r.Parent.Cells(r.Row, 2))) & _
                ", " & Trim(CStr(r.Parent.Cells(r.Row, 3))) & _
                ", " & Trim(CStr(r.Parent.Cells(r.Row, 4)))
                
                
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_main_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_cont_pnoc_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_del_conf_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_open_issues_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_order_release_status_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_osea_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_recent_build_plan_changes_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_resp_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_totals_sh_nm), CStr(strlink)
            iteracja_po_kolejnych_wierszach_dla CStr(SIXP.G_xq_sh_nm), CStr(strlink)
            '
            ''
            ' =========================================================
            ' =========================================================
        Else
            MsgBox "zadna akcja nie zostala podjeta"
            End
        End If
    Else
        MsgBox "dane z aktywnej komorki nie moga zostac usuniete, apka konczy dzialanie!"
        End
    End If
        
End Sub


Private Sub iteracja_po_kolejnych_wierszach_dla(shnm As String, strlink As String, Optional like_comparison As Boolean)


    ' strlink jako drugi argument jest zbiorem textowym z pierwszych 4 kolumn wybranej aktywnej komorki
    ' w celu zlinkowania potencjalnych danych do usuniecia
    
    ' od lutego 2017 powyzsza definicja zostala wzbogacona o tzw. wildcardy z niezaleznego formularza
    ' teraz strlink rowniez jest zapisem od usera ktory wybranym wyrazeniem chce usunac jego wszystkie wystapienia
    ' w pierwszym czterech kolumnach danych powiazanych
    
    ' aby bylo to wszystko mozliwe dorzucilem opcjonalny argument trzecu w postaci mozliwosci wyboru miedzy warunkiem if dokladnym
    ' : wykorzystanie znaku =
    ' oraz niedokoladnym, przyblizonym : wykorzystanie slowa kluczowego Like
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(shnm)
    
    On Error Resume Next
    sh.ShowAllData
    
    Dim r As Range
    Set r = sh.Range("A2")
    
    Do
        tmp = Trim(CStr(r.Parent.Cells(r.Row, 1))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 2))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 3))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 4)))
            
        If like_comparison = False Then
        
            If Trim(tmp) = Trim(strlink) Then
            
                r.EntireRow.Delete
            
                Set r = sh.Range("A1")
            End If
        ElseIf like_comparison = True Then
            If Trim(tmp) Like "*" & Trim(strlink) & "*" Then
            
                r.EntireRow.Delete
            
                Set r = sh.Range("A1")
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(CStr(r)) = ""
    
    
End Sub
