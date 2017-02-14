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
Private oi As Worksheet

Private r As Range

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
            
            .ComboBoxStatus.Clear
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putCross.Value)
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putTriangle.Value)
            .ComboBoxStatus.AddItem CStr(SIXP.GlobalCrossTriangleCircleModule.putCircle.Value)
            
            
            ' lista open issues
            .ListBox1.Clear

            
            ' ---------------------------------------------------------
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            .BtnDodajNowy.Enabled = True
            
            ' listbox otwartych issues w projekcie
            .ListBox1.Clear
            
            X = 1
            
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
                        If X = 1 Then
                            .TextBoxComment = CStr(r.Offset(0, SIXP.e_open_issues_comment - 1))
                            .TextBoxDelivery = CStr(r.Offset(0, SIXP.e_open_issues_delivery - 1))
                            .TextBoxNoOfPNs = CStr(r.Offset(0, SIXP.e_open_issues_no_of_pn - 1))
                            .TextBoxPartSupplier = CStr(r.Offset(0, SIXP.e_open_issues_part_supplier - 1))
                            .ComboBoxStatus = CStr(r.Offset(0, SIXP.e_open_issues_status - 1))
                        End If
                        ' ---------------------------------------------------------

                
                        ' Exit Do
                        .ListBox1.AddItem "Open issue #" & CStr(X) & ", " & r.Address
                        X = X + 1
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





