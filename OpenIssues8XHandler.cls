VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenIssues8XHandler"
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


' klasa ta ma sie zajmowac kontrola danych miedzy order releasami a main ogolnie
' dobrze by bylo potrfila wpisywac ladnie aktualny ycw dla update'u arkusza
' order release, dzieki czemu bedziemy mogli miec porownanie z aktualnym cw w arkuszu main

Private m As Worksheet
Private master As Worksheet
Private oi As Worksheet

Private r As Range


Private dunses As Dictionary

' klucz powrotu do arkuszy main oraz open issues (konkatenacja proj, plt, faza, cw chyba w innym orderze)
Public Label As String


Private link_zapamietaj As T_Link


Public currProj As String



Private Sub Class_Initialize()


    ' ten pseudo konstruktor nie bedzie mial narazie wiecej kodu
    ' 2016-08-10

    ' main sheets assigning
    ' ====================================================================
    ' przypisanie arkusza main
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' arkusz order release status handler
    Set oi = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set oi = Nothing
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormOpenIssues)


    currProj = proj


    If CStr(eca) = SIXP.G_BTN_TEXT_ADD Then
    
        With o
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            .BtnDodajNowy.Enabled = False
            
            
            
            'Public Enum E_RECENT_BP_CH
            '    e_recent_bp_ch_no_of_veh = 5
            '    e_recent_bp_ch_tbw
            '    e_recent_bp_ch_order_release_ch
            '    e_recent_bp_ch_comment
            'End Enum
            
            
            ' pola liczbowe
            ' ---------------------------------------------------------
            ' example
            ' .TextBoxAfterALTMRD = "0"
            .TextBoxComment = ""
            .TextBoxDelivery = ""
            .TextBoxNoOfPNs = ""
            .TextBoxPartSupplier = ""
            
            .CheckBoxVisible.Value = True
            
            .ComboBoxStatus.Clear
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putCross.Value)
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putTriangle.Value)
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putCircle.Value)
            
            
            ' lista open issues
            .ListBox1.Clear
            .ListBox1.MultiSelect = fmMultiSelectSingle
            
            ' ---------------------------------------------------------
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            .BtnDodajNowy.Enabled = True
            
            ' listbox otwartych issues w projekcie
            .ListBox1.Clear
            .ListBox1.MultiSelect = fmMultiSelectSingle
            
            x = 1
            
            Set r = ThisWorkbook.Sheets(SIXP.G_open_issues_sh_nm).Cells(1, 1)
            
            ' petla pobierania info z arkusza
            ' watpie ze kiedy kolwiek bardzie nagminne edytowania juz wpisanych issues zatem
            ' ogranicze sie do nazwania ich kolejnymi numerkami
            
            Do
                If CStr(.LabelTitle.Caption) = _
                    CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                

                        
                        ' pola liczbowe
                        ' ---------------------------------------------------------

                        ' example
                        ' .TextBoxAfterALTMRD = CStr(r.Offset(0, SIXP.e_del_conf_after_alt - 1))
                        If x = 1 Then
                            .TextBoxComment = CStr(r.Offset(0, SIXP.e_open_issues_comment - 1))
                            .TextBoxDelivery = CStr(r.Offset(0, SIXP.e_open_issues_delivery - 1))
                            .TextBoxNoOfPNs = CStr(r.Offset(0, SIXP.e_open_issues_no_of_pn - 1))
                            .TextBoxPartSupplier = CStr(r.Offset(0, SIXP.e_open_issues_part_supplier - 1))
                            .ComboBoxStatus = CStr(r.Offset(0, SIXP.e_open_issues_status - 1))
                            
                            If CStr(r.Offset(0, SIXP.e_open_issues_visible - 1)) = "1" Then
                                .CheckBoxVisible.Value = True
                            Else
                                .CheckBoxVisible.Value = False
                            End If
                        End If
                        ' ---------------------------------------------------------
                        
                        
                        ' ' Exit Do
                        ' Me.ListBox1.AddItem "Open issue #" & CStr(X) & ", " & _
                        ' CStr(Trim(r.Offset(0, SIXP.e_open_issues_part_supplier - 1))) & ", " & _
                        ' r.Address
                
                        ' Exit Do
                        ' .ListBox1.AddItem "Open issue #" & CStr(X) & ", " & r.Address
                        
                        .ListBox1.AddItem "Open issue #" & CStr(x) & ", " & _
                        CStr(Trim(r.Offset(0, SIXP.e_open_issues_part_supplier - 1))) & ", " & _
                        r.Address
                        
                        x = x + 1
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            .ListBox1.Selected(0) = True
            .ComboBoxStatus.Clear
            
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putCross.Value)
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putTriangle.Value)
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putCircle.Value)
        End With
    End If
End Sub


' trzeba by dodac sekcje odpowiadajaca za zarzadzanie mass importem oraz jako tako lista open issues
' bo aktualnie z CRUDa mamy tylko create, update - no i szczatkowe read, ktore jednak jest nieczytelne
'-----------------------------------------------------------------------------------------------------------------------------------

Public Function wypelnij_slownik_dunsami_jako_kluczami_pod_open_issues_z_mass_impotu_z_wizarda( _
    open_issues_sh As Worksheet, master_sh As Worksheet, mr As Range, details_sh As Worksheet) As Boolean


    Set master = master_sh
    ' mr = range a2 na arkuszu master wybranego wizarda
    
    
    
    Set io = open_issues_sh
    
    Set dunses = Nothing
    Set dunses = New Dictionary
    
    Dim item As OpenIssueItem
    Set r = mr
    Do
    
        If Trim(r.Offset(0, SIXP.Comments - 1)) <> "" Then
        
            If dc_is_red_or_yellow(details_sh, r.Offset(0, SIXP.Delivery_confirmation - 1)) Then
            
                ' logika dodawania, lub edycji nowego klucza w postaci # DUNS
                
                If dunses.Exists(r.Offset(0, SIXP.duns - 1)) Then
                
                    Set item = dunses(r.Offset(0, SIXP.duns - 1))
                    item.cmnt = item.cmnt & Trim(r.Offset(0, SIXP.Comments - 1)) & Chr(10) & Chr(10)
                    item.pns.Add CStr(r.Offset(0, SIXP.pn - 1))
                    
                    item.part_slash_supplier = Trim(r.Offset(0, SIXP.Supplier_Name - 1))
                    
                    
                Else
                
                    Set item = New OpenIssueItem
                    
                    ' trzeba budowac caly obiekt
                    item.duns = Trim(r.Offset(0, SIXP.duns - 1))
                    item.cmnt = Trim(r.Offset(0, SIXP.Comments - 1)) & Chr(10) & Chr(10)
                    
                    If Len(item.cmnt) > 10 Then
                        item.name = Trim(Left(item.cmnt, 10))
                    Else
                        item.name = item.cmnt
                    End If
                    
                    item.part_slash_supplier = Trim(r.Offset(0, SIXP.pn - 1))
                    Set item.pns = New Collection
                    item.pns.Add CStr(r.Offset(0, SIXP.pn - 1))
                    On Error Resume Next
                    item.status = SIXP.GlobalCrossTriangleCircleModule.putCross
                    item.widoczny = True
                    
                    
                    
                    
                    dunses.Add Trim(r.Offset(0, SIXP.duns - 1)), item
                End If
             End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until r.Row = 10000
    
    ' key as DUNS
    ' value as Object with custom data
    
    If dunses.Count > 0 Then
        wypelnij_slownik_dunsami_jako_kluczami_pod_open_issues_z_mass_impotu_z_wizarda = True
    Else
        wypelnij_slownik_dunsami_jako_kluczami_pod_open_issues_z_mass_impotu_z_wizarda = False
    End If
End Function

Private Function dc_is_red_or_yellow(details_sh, r As Range) As Boolean

    dc_is_red_or_yellow = False
    
    Dim mrd_r As String
    
    ' komorka zawierajaca Y????CW?? MRD
    ' ====================================
    mrd_r = CStr(details_sh.Range("B9"))
    ' ====================================
    
    
    
    If (Trim(r) = "") Or (Trim(r) Like "*OPEN*") Or (Trim(r) Like "*POT*ITDC*") Or (Trim(r) Like "*N/A*") Then
        dc_is_red_or_yellow = True
    ElseIf Trim(r) Like "*Y*CW*" Then
        ' sprawdz, czy to jest opozniony
        
        gdzie_jest_y = Application.WorksheetFunction.Find("Y", CStr(r))
        dc_liczba = CLng(CStr(Mid(r, gdzie_jest_y + 1, 4)) & CStr(Right(r, 2)))
        
        ' liczba naszee mrd
        mrd_liczba = CLng(CStr(Mid(mrd_r, 2, 4)) & CStr(Right(mrd_r, 2)))
        
        If CLng(dc_liczba) > CLng(mrd_liczba) Then
            dc_is_red_or_yellow = True
        End If
    End If
    
    
End Function


Public Sub zrzuc_dane_ze_slownika_do_arkusza_open_issues()

    ' slownik zawiera wszystko to, co bylo w arkuszu master wybranego wizarda
    ' ==========================================================================
    Dim item As OpenIssueItem
    Set r = znajdz_pierwszy_wolny_wiersz_w_open_issues()
    
    
    ' przyklad porownania labelki do zawartosci arkusza
    ' If CStr(Me.LabelTitle.Caption) = _
    '   CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
    
    arr_labela = Split(Me.currProj, ", ")
    
    For Each Key In dunses
    
    
        
        For i = LBound(arr_labela) To UBound(arr_labela)
        
        
        
        
            j = i
            
            If LBound(arr_labela) = 0 Then
                j = i
            
            ElseIf LBound(arr_labela) > 0 Then
                j = i - LBound(arr_labela)
                
            Else
                MsgBox "poszlo cos nie tak!"
                End
            End If
            
            ' j zostal odpowiednio zmodyfikowany
            r.Offset(0, j).Value = CStr(arr_labela(i))
        Next i
        
        
        
        Set link_zapamietaj = New T_Link
        link_zapamietaj.zrob_mnie_z_range r.Parent.Cells(r.Row, 1)
        
        Set item = dunses(Key)
        r.Offset(0, SIXP.e_open_issues_status - 1) = item.status
        r.Offset(0, SIXP.e_open_issues_no_of_pn - 1) = item.pns.Count
        r.Offset(0, SIXP.e_open_issues_part_supplier - 1) = item.part_slash_supplier
        r.Offset(0, SIXP.e_open_issues_delivery - 1) = item.delivery_date
        r.Offset(0, SIXP.e_open_issues_comment - 1) = item.cmnt
        
        If item.widoczny Then
            r.Offset(0, SIXP.e_open_issues_visible - 1) = "1"
        Else
            r.Offset(0, SIXP.e_open_issues_visible - 1) = "0"
        End If
        
        
        Set r = r.Offset(1, 0)
        
    Next
    
    
    ' ==========================================================================
End Sub


Private Function znajdz_pierwszy_wolny_wiersz_w_open_issues() As Range

    Dim tmp As Range
    Set tmp = oi.Range("A1")
    Do
    
        Set tmp = tmp.Offset(1, 0)
    Loop Until Trim(tmp) = ""
    
    Set znajdz_pierwszy_wolny_wiersz_w_open_issues = tmp
    
End Function

Public Sub usun_duplikaty()
    
    ' usuwamy duplikaty z oi sheet
    
    ' przyklad porownania labelki do zawartosci arkusza
    ' If CStr(Me.LabelTitle.Caption) = _
    '   CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then

    
    If Not oi Is Nothing Then
        ' =========================
        ''
        '
        ' kolejny raz heurystycznie patrze grubo od dolu gdyby znowu sie okazalo ze ktos sobie zostawia "wolne pola"
        oi.Range(Cells(1, 1), Cells(SIXP.e_open_issues_visible, oi.Range("A100000").End(xlUp))).RemoveDuplicates Array(SIXP.e_link_project, _
            SIXP.e_link_plt, _
            SIXP.e_link_faza, _
            SIXP.e_link_cw, _
            SIXP.e_open_issues_status, _
            SIXP.e_open_issues_no_of_pn, _
            SIXP.e_open_issues_part_supplier, _
            SIXP.e_open_issues_delivery, _
            SIXP.e_open_issues_comment), _
            xlYes
        '
        ''
        ' =========================
    End If
End Sub


Public Sub otworz_formularz_ponownie()
    
    ' otwiermy bezposrednio form dla arkusza open issues
    ' -----------------------------------------------------
    Dim l As New Linker
    txt_l = l.return_full_concated_r_string_comma_seperated(link_zapamietaj)
    
    
    SIXP.DataFlowPodFormMainModule.zrob_open_issues CStr("Edytuj"), CStr(txt_l)
    ' -----------------------------------------------------
End Sub

Public Sub otworz_formularz_ponownie_bez_danych(txt)
    MsgBox "brak danych!"
    SIXP.DataFlowPodFormMainModule.zrob_open_issues SIXP.G_BTN_TEXT_ADD, CStr(txt)
End Sub


Public Sub oznacz_wlasciwy_rekord_w_arkuszu_main_po_mass_imporcie()


    ' przyklad porownania labelki do zawartosci arkusza
    ' If CStr(Me.LabelTitle.Caption) = _
    '   CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
    
    

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------





