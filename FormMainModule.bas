Attribute VB_Name = "FormMainModule"
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

Public Sub run_FormMain(Optional link_str As String)



    If Trim(link_str) = "" Then
    
        ' MsgBox "Nie ma danych do zlinkowania!"
        ' jesli nie ma jeszcze postawionych danych
        ' sprobujmy uzupelnic link na nowo
        
        Dim l As T_Link
        Dim checker As SixPChecker
        Set checker = New SixPChecker
        
        If checker.sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main() Then
            If Trim(ActiveCell.Parent.Cells(ActiveCell.Row, 1).Value) <> "" Then
                If ActiveCell.Row > 1 Then
                    Set l = New T_Link
                    Dim lr As Linker
                    Set lr = New Linker
                    l.zrob_mnie_z_argsow _
                        ActiveCell.Parent.Cells(ActiveCell.Row, 1), ActiveCell.Parent.Cells(ActiveCell.Row, 2), _
                        ActiveCell.Parent.Cells(ActiveCell.Row, 3), ActiveCell.Parent.Cells(ActiveCell.Row, 4)
                        
                    link_str = CStr(lr.return_full_concated_r_string_comma_seperated(l))
                End If
            End If
        End If
        
        If Trim(link_str) = "" Then
            ' sprawdzam drugi raz!
            MsgBox "Nie ma danych do zlinkowania!"
            Exit Sub
        End If
    End If


    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 1

    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.init link_str
    
    Set fmh = Nothing
    
    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 0
End Sub

Public Sub add_new_project()

    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 1
    
    Dim fmh As FormMainHandler
    Set fmh = New FormMainHandler
    
    fmh.new_project
    
    Set fmh = Nothing
    
    ThisWorkbook.Sheets(SIXP.G_register_sh_nm).Range("RUN") = 0
End Sub







Public Sub adjustuj_guzik(e_main As E_MAIN_ORDER, ish As Worksheet, imainsh As Worksheet, il As T_Link, paleta As PaletaTheDailyCommute)



    ' dzieki koncepcji "porownania posredniego" mamy latwosc z ogarnieciem
    ' jaki kolor dany button powinien miec
    ' wydzielilem podwojnie logike a w sumie cale dzialo te funkcji mozna byloby rozdzielic
    ' od razu na wysokosci suba ktory sie znajduje powyzej
    ' byloby nieco mniej kodu, ale skoro juz to napisalem
    ' to niech to tak zostanie - szkoda czasu i nie bede sobie marnowal statsow
    
    
    Dim range_from_main As Range, range_from_ish As Range
    
    ' porownanie posrednie
    Set range_from_main = il.znajdz_siebie_w_arkuszu(imainsh)
    Set range_from_ish = il.znajdz_siebie_w_arkuszu(ish)
    
    ' jesli te dwie zmienne nie sa puste to znaczy ze mamy takie dane spasowane i chcemy je edytowac
    If Not range_from_main Is Nothing Then
        If Not range_from_ish Is Nothing Then
            
            If e_main = e_main_last_update_on_order_release_status Then
            
                With FormMain.BtnOrderReleaseStatus
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_recent_build_plan_changes Then
                
                With FormMain.BtnRecentBuildPlanChanges
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_chart_contracted_pnoc Then
                
                With FormMain.BtnContractedPNOC
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_osea Then
                
                With FormMain.BtnOseaScope
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_totals Then
                
                With FormMain.BtnTotals
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_xq Then
                
                With FormMain.BtnXq
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_del_conf Then
                
                With FormMain.BtnDelConf
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
                
                
                With FormMain.BtnNewDelConf
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
                
            ElseIf e_main = e_main_last_update_on_open_issues Then
                
                With FormMain.BtnOpenIssues
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            ElseIf e_main = e_main_last_update_on_resp Then
                
                With FormMain.BtnResp
                    .Caption = SIXP.G_BTN_TEXT_EDIT
                    .BackColor = paleta.dark_grey
                    .ForeColor = paleta.orange
                End With
            End If
        Else
            ' sekcja, gdzie cos znalezlismy w main jednak nie ma tego w arkuszu przeszukiwanym
            ' to znaczy tyle ze trzeba dodac nowe
            
            If e_main = e_main_last_update_on_order_release_status Then
            
                FormMain.BtnOrderReleaseStatus.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOrderReleaseStatus.BackColor = paleta.yellow
                FormMain.BtnOrderReleaseStatus.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_recent_build_plan_changes Then
            
                FormMain.BtnRecentBuildPlanChanges.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnRecentBuildPlanChanges.BackColor = paleta.yellow
                FormMain.BtnRecentBuildPlanChanges.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_chart_contracted_pnoc Then
            
                FormMain.BtnContractedPNOC.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnContractedPNOC.BackColor = paleta.yellow
                FormMain.BtnContractedPNOC.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_osea Then
            
                FormMain.BtnOseaScope.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOseaScope.BackColor = paleta.yellow
                FormMain.BtnOseaScope.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_totals Then
            
                FormMain.BtnTotals.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnTotals.BackColor = paleta.yellow
                FormMain.BtnTotals.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_xq Then
            
                FormMain.BtnXq.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnXq.BackColor = paleta.yellow
                FormMain.BtnXq.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_del_conf Then
            
                FormMain.BtnDelConf.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnDelConf.BackColor = paleta.yellow
                FormMain.BtnDelConf.ForeColor = paleta.dark_grey
                
                
                FormMain.BtnNewDelConf.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnNewDelConf.BackColor = paleta.yellow
                FormMain.BtnNewDelConf.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_open_issues Then
            
                FormMain.BtnOpenIssues.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnOpenIssues.BackColor = paleta.yellow
                FormMain.BtnOpenIssues.ForeColor = paleta.dark_grey
                
            ElseIf e_main = e_main_last_update_on_resp Then
            
                FormMain.BtnResp.Caption = SIXP.G_BTN_TEXT_ADD
                FormMain.BtnResp.BackColor = paleta.yellow
                FormMain.BtnResp.ForeColor = paleta.dark_grey
                
            End If
        End If
    End If
    
    
    
    
End Sub


