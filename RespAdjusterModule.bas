Attribute VB_Name = "RespAdjusterModule"
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


Public Sub resp_adjuster()
    
    ' this sub opens worksheet which helps to provide if keyword is under FMA resp
    ' ------------------------------------------------------------------------------
    ' ------------------------------------------------------------------------------
    
    FormRespAdjuster.ListBoxInScope.Clear
    FormRespAdjuster.ListBoxOutOfScope.Clear
    
    Dim d As Dictionary
    Set d = verify_with_wizard_buff()
    Set d = complete_data_from_register(d, ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2"))
    go_to_register_and_prepare_form d, ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("G2")
    
    SIXP.FormRespAdjuster.Show
End Sub

Private Sub go_to_register_and_prepare_form(ByRef dd As Dictionary, rf As Range)

    ' slownik pobrany jako argument zawiera elementy nazw resp (m. in. FMA, OSEA, SC, MPC)
    ' --------------------------------------------------------------------------------------
    ''
    '
    
    ' nie chce laczyc sie miedzy subami i funkcjami
    ' jednak zalozenie jest takie ze funkcje wczesniejsze przygotowaly slownik
    ' tak aby dla keyi byly trzy odpowiadajace wartosci 0,1,2
    ' gdzie:
    
    ' 1 - dane sa dostepne tylko na wiz buff (potencjalnie nowe wartosci)
    ' 2 - dane sa dostepne tylko w rejestrze
    
    ' pierwszy pusty
    If Trim(rf.Offset(1, 0)) <> "" Then
        Set rf = rf.End(xlDown).Offset(1, 0)
    Else
        Set rf = rf.Offset(1, 0)
    End If
    
    For Each Key In dd
        If dd(Key) = 1 Then
            rf = Key
            rf.Offset(0, 1) = 0
            
            Set rf = rf.Offset(1, 0)
        End If
    Next
    
    
    With SIXP.FormRespAdjuster
    
        Set rf = rf.End(xlUp).End(xlUp).Offset(1, 0)
        Do
            
            If Trim(rf.Offset(0, 1)) = "1" Then
                .ListBoxInScope.AddItem Trim(rf)
            Else
                .ListBoxOutOfScope.AddItem Trim(rf)
            End If
            Set rf = rf.Offset(1, 0)
        Loop Until Trim(rf) = ""
    End With
    
    
    
    '
    ''
    ' --------------------------------------------------------------------------------------
    
End Sub


Private Function verify_with_wizard_buff()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    
    Dim r As Range
    Set r = sh.Range("B2")
    
    
    Dim d As Dictionary
    ' r jako referencja
    ' drugi i trzeci argument jako parametry offsety - czy lecimy po kolumnach, czy lecimy po wierszach
    ' ostatni, czwarty parametr sluzy wartosci wpisanej dla nowego elementu slownika
    ' i dla przykladu jesli wpada nowy element z wizard buff przypisujemy mu wartosc = 1
    Set d = wypelnij_dic(r, 0, 1, 1)
    
    
    Set verify_with_wizard_buff = d
End Function

Private Function wypelnij_dic(rr As Range, w, k, val) As Dictionary
    Set wypelnij_dic = Nothing
    
    Dim dic_tmp As Dictionary
    
    If Trim(rr) <> "" Then
        Set dic_tmp = New Dictionary
        
        Do
            If Not dic_tmp.Exists(Trim(rr)) Then
                dic_tmp.Add Trim(rr), val
            Else
                dic_tmp(Trim(rr)) = val
            End If
            Set rr = rr.Offset(w, k)
        Loop Until Trim(rr) = ""
        
        Set wypelnij_dic = dic_tmp
        Exit Function
    End If
End Function

Private Function complete_data_from_register(dd As Dictionary, rr As Range) As Dictionary


    ' dd to slownik z juz czesciowo skomplemtowanymi danymi z wizard buff
    ' rr poczatek zasiegu skomplemtowanych danych w rejestrze
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(SIXP.G_register_sh_nm)
    
    Dim r As Range
    Set r = sh.Range("G2")
    
    
    Dim d As Dictionary
    ' r jako referencja
    ' drugi i trzeci argument jako parametry offsety - czy lecimy po kolumnach, czy lecimy po wierszach
    ' ostatni, czwarty parametr sluzy wartosci wpisanej dla nowego elementu slownika
    ' i dla przykladu jesli wpada nowy element z rejestru przypisujemy mu wartosc = 2
    Set d = wypelnij_dic(r, 1, 0, 2)
    
    For Each Key In dd
        If Not d.Exists(Key) Then
            d.Add Key, dd(Key)
        End If
    Next
    
    
    Set complete_data_from_register = d
    
    
End Function




' =================================================================
Public Function podlicz_w_zgodzie_z_ukladem_z_arkusza_register() As Long
    podlicz_w_zgodzie_z_ukladem_z_arkusza_register = 0
    
    Dim wiz_buff As Worksheet
    Set wiz_buff = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)
    Dim reg_sh As Worksheet
    Set reg_sh = ThisWorkbook.Sheets(SIXP.G_register_sh_nm)
    
    Dim r As Range, r_reg As Range
    ' poczatek listy w rejestrze
    
    Set r = wiz_buff.Cells(3, 2)
    Do
        ' ----------------------------------
        Set r_reg = reg_sh.Range("G2")
        Do
            If Trim(r.Offset(-1, 0)) = Trim(r_reg) Then
                If CStr(r_reg.Offset(0, 1)) = "1" Then
                
                    If IsNumeric(r) Then
                        podlicz_w_zgodzie_z_ukladem_z_arkusza_register = podlicz_w_zgodzie_z_ukladem_z_arkusza_register + CLng(r)
                        Exit Do
                    Else
                        MsgBox "podlicz_w_zgodzie_z_ukladem_z_arkusza_register: NaN error!"
                        End
                    End If
                End If
            End If
            
            
            Set r_reg = r_reg.Offset(1, 0)
        Loop Until Trim(r_reg) = ""
        ' ----------------------------------
        Set r = r.Offset(0, 1)
    Loop Until Trim(r) = ""
End Function



Public Sub test_on_g1_h1()

    Dim wrksh As Worksheet
    Set wrksh = ThisWorkbook.Sheets(SIXP.G_WIZARD_BUFF_SH_NM)

    SIXP.RespAdjusterModule.resp_adjuster
    wrksh.Range("G1").Value = "IN SCOPE"
    wrksh.Range("H1").Value = CStr(podlicz_w_zgodzie_z_ukladem_z_arkusza_register())
End Sub
