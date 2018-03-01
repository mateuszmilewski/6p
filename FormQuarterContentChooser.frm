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


Private Sub BtnClear_Click()


    For X = 0 To Me.ListBoxProj.ListCount - 1

        Me.ListBoxProj.Selected(X) = False
        Me.ListBoxCW.Selected(X) = False
        Me.ListBoxFaza.Selected(X) = False
        Me.ListBoxPLT.Selected(X) = False
        
    Next X
End Sub


' MAIN SUB TO GET DATA FROM QUARTER!!!
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------
Private Sub BtnImport_Click()
    
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
            For X = 0 To Me.ListBoxRow.ListCount - 1
                c.Add Me.ListBoxRow.List(X)
            Next X
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
    
    For X = 0 To Me.ListBoxRow.ListCount - 1
        If Me.ListBoxRow.Selected(X) Then
        
            checkIfThereIsAnySelection = True
        
            c.Add Me.ListBoxRow.List(X)
        End If
    Next X

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
    
    For X = 0 To Me.ListBoxProj.ListCount - 1
        If Me.ListBoxProj.Selected(X) Then
            
            
            Me.ListBoxCW.Selected(X) = True
            Me.ListBoxFaza.Selected(X) = True
            Me.ListBoxPLT.Selected(X) = True
            
            Me.ListBoxRow.Selected(X) = True
            
        Else
            
            Me.ListBoxCW.Selected(X) = False
            Me.ListBoxFaza.Selected(X) = False
            Me.ListBoxPLT.Selected(X) = False
            
            Me.ListBoxRow.Selected(X) = False
            
        End If
    Next X
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
    
    For X = LBound(iarr) To UBound(iarr)
        If X = LBound(iarr) Then
            
            ' PROJECT section
            iproj = iarr(X)
        End If
        
        If X = LBound(iarr) + 1 Then
            iplt = iarr(X)
        End If
        
        If X = LBound(iarr) + 2 Then
            ifaza = iarr(X)
        End If
        
        If X = LBound(iarr) + 3 Then
            icw = iarr(X)
        End If
    Next X
    
    
    
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

