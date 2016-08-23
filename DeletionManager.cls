VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeletionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private Sub iteracja_po_kolejnych_wierszach_dla(shnm As String, strlink As String)


    ' strlink jako drugi argument jest zbiorem textowym z pierwszych 4 kolumn wybranej aktywnej komorki
    ' w celu zlinkowania potencjalnych danych do usuniecia
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(shnm)
    
    Dim r As Range
    Set r = sh.Range("A2")
    
    Do
        tmp = Trim(CStr(r.Parent.Cells(r.Row, 1))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 2))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 3))) & _
            ", " & Trim(CStr(r.Parent.Cells(r.Row, 4)))
            
        
        If Trim(tmp) = Trim(strlink) Then
        
            r.EntireRow.Delete
        
            Set r = sh.Range("A1")
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(CStr(r)) = ""
    
    
End Sub
