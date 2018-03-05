VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormQuarterContentChooser 
   Caption         =   "Chooser"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10170
   OleObjectBlob   =   "FormQuarterContentChooser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormQuarterContentChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' passed Dictionary
Private pD As Dictionary
Private qh As QuarterHandler
Private nazwaPliku As String


Public Sub passData(ByRef d As Dictionary, mqh As QuarterHandler, m_nazwaPliku)

    Set pD = d
    Set qh = mqh
    nazwaPliku = CStr(m_nazwaPliku)
    
End Sub


Private Sub BtnAll_Click()


    For x = 0 To Me.ListBoxProj.ListCount - 1

        Me.ListBoxProj.Selected(x) = True
        Me.ListBoxCW.Selected(x) = True
        Me.ListBoxFaza.Selected(x) = True
        Me.ListBoxPLT.Selected(x) = True
        
    Next x

End Sub

Private Sub BtnClear_Click()


    For x = 0 To Me.ListBoxProj.ListCount - 1

        Me.ListBoxProj.Selected(x) = False
        Me.ListBoxCW.Selected(x) = False
        Me.ListBoxFaza.Selected(x) = False
        Me.ListBoxPLT.Selected(x) = False
        
    Next x
End Sub


' MAIN SUB TO GET DATA FROM QUARTER!!!
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
Private Sub BtnImport_Click()
    
    
    SIXP.GlobalFooModule.gotoThisWorkbookMainA1
    
    Hide


    ' polacz z powrotem z arkuszem main_table z Quartera
    
    Dim mt As Worksheet
    Set mt = tryAssignMainTableFromQuarter(qh, nazwaPliku)
    
    
    If Not mt Is Nothing Then
    
    
        ' kolekcja wierszy (liczba bez zawartosci komorek)
        Dim c As Collection
        Set c = New Collection
        
        If checkIfThereIsAnySelection(c) Then
            ' bier tylko selekcje - co sie stalo juz w "stejtmancie" - ot taka oszczednosc!
        Else
            ' bier wszystko
            For x = 0 To Me.ListBoxRow.ListCount - 1
                c.Add Me.ListBoxRow.List(x)
            Next x
        End If
    
        If Not qh Is Nothing Then
            qh.copyDataFromQuarter mt, c
        Else
            ' nie ma qh, zatem trzeba uruchomic na nowo
            Set qh = New QuarterHandler
            qh.copyDataFromQuarter mt, c
        End If
    Else
        MsgBox "utracono polaczenie z quarterem!"
        End
    End If

End Sub
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------

Public Function checkIfThereIsAnySelection(c As Collection) As Boolean

    checkIfThereIsAnySelection = False
    
    For x = 0 To Me.ListBoxRow.ListCount - 1
        If Me.ListBoxRow.Selected(x) Then
        
            checkIfThereIsAnySelection = True
        
            c.Add Me.ListBoxRow.List(x)
        End If
    Next x

End Function


Private Function tryAssignMainTableFromQuarter(qh As QuarterHandler, np) As Worksheet

    Set tryAssignMainTableFromQuarter = Nothing
    
    
    Dim tmp As Worksheet
    Set tmp = Nothing
    
    If Not qh Is Nothing Then
        Set tmp = qh.getQuarterMainTable()
    Else
        Set tmp = Workbooks(CStr(np)).Sheets(SIXP.G_MAIN_TB_FROM_Q)
    End If
    
    Set tryAssignMainTableFromQuarter = tmp
    
End Function


Private Sub BtnReset_Click()

    If Not pD Is Nothing Then
        If pD.Count > 0 And (Not qh Is Nothing) Then
            
            ' -----------------------------------------------------------
            qh.overwriteDictionary pD
            qh.resetListBoxes
            
            ' -----------------------------------------------------------
        
        Else
            MsgBox "nie mozna zeresetowac, cos poszlo nie tak!"
        End If
    
    Else
        MsgBox "nie mozna zeresetowac, cos poszlo nie tak!"
        
    End If
    
End Sub

Private Sub ListBoxProj_Change()
    inner_listbox_action
End Sub

Private Sub ListBoxProj_Click()
    inner_listbox_action
End Sub


Private Sub inner_listbox_action()
    
    For x = 0 To Me.ListBoxProj.ListCount - 1
        If Me.ListBoxProj.Selected(x) Then
            
            
            Me.ListBoxCW.Selected(x) = True
            Me.ListBoxFaza.Selected(x) = True
            Me.ListBoxPLT.Selected(x) = True
            
            Me.ListBoxRow.Selected(x) = True
            
        Else
            
            Me.ListBoxCW.Selected(x) = False
            Me.ListBoxFaza.Selected(x) = False
            Me.ListBoxPLT.Selected(x) = False
            
            Me.ListBoxRow.Selected(x) = False
            
        End If
    Next x
End Sub

Private Sub TextBox1_Change()
    innerWildcardLogic
End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    innerWildcardLogic
End Sub

Private Sub TextBox1_Enter()
    innerWildcardLogic
End Sub


Private Sub innerWildcardLogic()

    iarr = Split(Me.TextBox1.Value, ";")


    iproj = ""
    iplt = ""
    ifaza = ""
    icw = ""
    
    For x = LBound(iarr) To UBound(iarr)
        If x = LBound(iarr) Then
            
            ' PROJECT section
            iproj = iarr(x)
        End If
        
        If x = LBound(iarr) + 1 Then
            iplt = iarr(x)
        End If
        
        If x = LBound(iarr) + 2 Then
            ifaza = iarr(x)
        End If
        
        If x = LBound(iarr) + 3 Then
            icw = iarr(x)
        End If
    Next x
    
    
    
    If Not pD Is Nothing Then
        If pD.Count > 0 And (Not qh Is Nothing) Then
            
            ' -----------------------------------------------------------
            qh.overwriteDictionary pD
            qh.wildcardListBoxes iproj, iplt, ifaza, icw
            
            ' -----------------------------------------------------------
        
        Else
            MsgBox "nie mozna uzyc wildcard, cos poszlo nie tak!"
        End If
    
    Else
        MsgBox "nie mozna uzyc wildcard, cos poszlo nie tak!"
        
    End If
    
    
    
End Sub

Private Sub UserForm_Activate()
     If Not qh Is Nothing Then
        qh.noSelectionOnListBoxes
     End If
End Sub

