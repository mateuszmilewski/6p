VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormCatchQuarter 
   Caption         =   "Q"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FormCatchQuarter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormCatchQuarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBoxFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
            
            
    SIXP.GlobalFooModule.gotoThisWorkbookMainA1
            
    Hide
    
    Dim w As Workbook
    Dim Sh As Worksheet
    Set Sh = Nothing
    
    
    For Each w In Workbooks
        
        If w.name = Me.ListBoxFiles.Value Then
            On Error Resume Next
            Set Sh = w.Sheets(SIXP.G_MAIN_TB_FROM_Q)
            Exit For
        End If
    Next w
    
    
    If Not Sh Is Nothing Then
    
    
        If checkFields(Sh) Then

            ' sub znajduje sie w modle Quarter
            uruchomLogikePrzechwytywaniaDanychZStaregoWybranegoQuartera Sh, w.name
        Else
            MsgBox "arkusz: " & SIXP.G_MAIN_TB_FROM_Q & " nie posiada prawidlowo ustawionych kolumn!"
            End
        End If
        
    
    Else
        MsgBox "ten plik nie posiada arkusza: " & SIXP.G_MAIN_TB_FROM_Q & "!"
        End
    End If
    
    
End Sub


Private Function checkFields(ByRef Sh As Worksheet) As Boolean

    checkFields = False
    
    
    ' nie jest to dokladne sprawdzenie, a jedynie szczatkowe sprawdzenie pierwszych 6 kolumn arkusza main table
    ' wydaje mi sie jednak, ze takie sprawdzenie w wiekszosci zalatwia sprawe i mozna by prawie pewnym, ze mamy do
    ' czynienia z plikiem typu quarter - gdyby omylkowo ktos wybral inny plik prawdopodobienstwo, ze akurat wybierze
    ' plik inny ktory ma taki sam uklad danych jest raczej male
    
    If Sh.Cells(1, 1).Value = "Project" Then
        If Sh.Cells(1, 2).Value = "Plant" Then
            If Sh.Cells(1, 3).Value = "PHASE" Then
                If Sh.Cells(1, 4).Value = "CW" Then
                    If Sh.Cells(1, 5).Value = "STATUS" Then
                        If Sh.Cells(1, 6).Value = "MRD" Then
                        
                            checkFields = True
                            
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
