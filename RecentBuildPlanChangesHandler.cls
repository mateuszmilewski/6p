VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecentBuildPlanChangesHandler"
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
Private rsh As Worksheet

Private r As Range

Private Sub Class_Initialize()


    ' ten pseudo konstruktor nie bedzie mial narazie wiecej kodu
    ' 2016-08-10

    ' main sheets assigning
    ' ====================================================================
    ' przypisanie arkusza main
    Set m = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    ' arkusz order release status handler
    Set rsh = ThisWorkbook.Sheets(SIXP.G_recent_build_plan_changes_sh_nm)
    ' ====================================================================
    
    

End Sub

Private Sub Class_Terminate()
    Set m = Nothing
    Set rsh = Nothing
End Sub


Public Sub inicjuj_wstepne_wartosci_pol_do_wypelnienia(proj As String, eca As String, o As FormRecentBuildPlanChanges)


    If CStr(eca) = SIXP.G_BTN_TEXT_ADD Then
    
        With o
            
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            
            'Public Enum E_RECENT_BP_CH
            '    e_recent_bp_ch_no_of_veh = 5
            '    e_recent_bp_ch_tbw
            '    e_recent_bp_ch_order_release_ch
            '    e_recent_bp_ch_comment
            'End Enum
            
            ' pola powiazane datami
            ' ---------------------------------------------------------
            .DTPickerReleased = Date
            .TextBoxReleased = CStr(parse_from_date_to_yyyycw(Date))
            .DTPickerTBW = Date
            .TextBoxTBW = CStr(parse_from_date_to_yyyycw(Date))

            
            
            ' ---------------------------------------------------------
            
            ' pola liczbowe
            ' ---------------------------------------------------------
            .TextBoxNoOfVeh = "0"
            
            ' ---------------------------------------------------------
            
            ' zwykly text po prostu puste
            ' ---------------------------------------------------------
            .TextBoxCmnt = ""
            ' ---------------------------------------------------------
        End With
    
    ElseIf CStr(eca) = SIXP.G_BTN_TEXT_EDIT Then
    
        With o
    
            .LabelTitle.Caption = CStr(proj)
            .BtnSubmit.Caption = CStr(eca)
            
            Set r = rsh.Cells(1, 1)
            Do
                If CStr(.LabelTitle.Caption) = _
                    CStr(Trim(r) & ", " & Trim(r.Offset(0, 1)) & ", " & Trim(r.Offset(0, 2)) & ", " & Trim(r.Offset(0, 3))) Then
                
                
                
                        ' znalazlem odpowiednia linijke pora wypelnic fieldy od forma
                        ' ---------------------------------------------------------
                        ' pola powiazane datami
                        ' ---------------------------------------------------------
                        ' edit
                        .TextBoxReleased = CStr(r.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1))
                        .DTPickerReleased = SIXP.GlobalFooModule.from_yyyy_cw_to_monday_from_this_week(CStr(r.Offset(0, SIXP.e_recent_bp_ch_order_release_ch - 1)))

                        .TextBoxTBW = CStr(r.Offset(0, SIXP.e_recent_bp_ch_tbw - 1))
                        .DTPickerTBW = SIXP.GlobalFooModule.from_yyyy_cw_to_monday_from_this_week(CStr(r.Offset(0, SIXP.e_recent_bp_ch_tbw - 1)))

                        
                        ' ---------------------------------------------------------
                        
                        ' pola liczbowe
                        ' ---------------------------------------------------------

                        .TextBoxNoOfVeh = CStr(r.Offset(0, SIXP.e_recent_bp_ch_no_of_veh - 1))
                        
                        ' textowe po prostu
                        ' ---------------------------------------------------------
                        .TextBoxCmnt = CStr(r.Offset(0, SIXP.e_recent_bp_ch_comment - 1))
                        ' ---------------------------------------------------------
                    
                
                
                        Exit Do
                End If
                Set r = r.Offset(1, 0)
            Loop Until Trim(r) = ""
            
            
            
        
        End With
    End If
End Sub


