VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormGetExtractFromSechel 
   Caption         =   "Get Sechel Extract"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "FormGetExtractFromSechel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormGetExtractFromSechel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnParse_Click()
    Hide
    
    
    
    ' 3 arg: pattern po projekcie z exportu
    ' 4 arg: pattern po magazynie z exportu
    innerParse CStr(Me.ListBox1.Value), CDate(Me.DTPicker1.Value), CStr(Me.TextBoxSousprojet.Value), CStr(Me.TextBoxMag.Value)
End Sub



Private Sub innerParse(filename As String, d As Date, sousProjet As String, mag As String)



    If checkThisFile(filename) Then
        
        ' sprawdzone kilka kolumn pliku - wychodzi na to ze to jest to
        '
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Dim sb As Worksheet, src As Worksheet
        Set sb = ThisWorkbook.Sheets(SIXP.G_SECHEL_BUFF_SH_NM)
        Set src = Application.Workbooks(filename).ActiveSheet
        
        sb.Range("A1:ZZ1").EntireColumn.Clear
        
        Dim sideRangeCollection As Collection
        Set sideRangeCollection = New Collection
        
        sb.Range("A1").Value = "LINES"
        sb.Range("B1").Value = countLines(src, d, sousProjet, mag, sideRangeCollection)
        
        sb.Range("A2").Value = "RECU"
        sb.Range("B2").Value = countLikeSituation(sideRangeCollection, "re?u")
        
        ' FauxManquant
        sb.Range("A3").Value = "FauxManquant"
        sb.Range("B3").Value = countSituation(sideRangeCollection, "FauxManquant")
        
        
        'manquantPlus
        sb.Range("A4").Value = "manquantPlus"
        sb.Range("B4").Value = countSituation(sideRangeCollection, "manquantPlus")
        
        
        sb.Range("A5").Value = "A venir"
        sb.Range("B5").Value = countSituation(sideRangeCollection, "A venir")
        
        sb.Range("A6").Value = "en cours"
        sb.Range("B6").Value = countSituation(sideRangeCollection, "en cours")
        
        sb.Range("A7").Value = "manquant"
        sb.Range("B7").Value = countSituation(sideRangeCollection, "manquant")
        
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        '
        
    Else
        MsgBox "plik nie wyglada jakby byl exportem z Sechela...."
    End If

End Sub


Private Function countLikeSituation(c As Collection, ptrn As String)


    countLikeSituation = 0
    
    Dim r As Range
    For Each r In c
        
        If CStr(r.Offset(0, SIXP.E_SECHEL__Situation - 1).Value) Like CStr(ptrn) Then
            countLikeSituation = countLikeSituation + 1
        End If
    Next r
End Function


Private Function countSituation(c As Collection, ptrn As String)


    countSituation = 0
    
    Dim r As Range
    For Each r In c
        
        If CStr(r.Offset(0, SIXP.E_SECHEL__Situation - 1).Value) = CStr(ptrn) Then
            countSituation = countSituation + 1
        End If
    Next r
End Function


Private Function countLines(src As Worksheet, d As Date, sousProjet, mag, ByRef c As Collection)


    countLines = 0

    Dim r As Range
    Set r = src.Range("A2")
    
    Do
    
        If CDate(d) = CDate(r.Offset(0, SIXP.E_SECHEL__Date_echeance - 1).Value) Then
            If CStr(r.Offset(0, SIXP.E_SECHEL__Sousprojet - 1).Value) Like "*" & CStr(sousProjet) & "*" Then
                If CStr(mag) = CStr(r.Offset(0, SIXP.E_SECHEL__Mag - 1).Value) Then
                    c.Add r
                End If
            End If
        End If
        
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
    
    
    countLines = c.Count
    
End Function


Private Function checkThisFile(fn As String) As Boolean

    checkThisFile = False
    
    Dim sh As Worksheet
    Set sh = Application.Workbooks(fn).ActiveSheet
    
    
    If sh.Cells(1, 1).Value = "GAc_Nom_NOA" Then
        If sh.Cells(1, 2).Value = "Article" Then
            If sh.Cells(1, 5).Value = "Nom_fournisseur" Then
            
                checkThisFile = True
            End If
        End If
    End If

End Function






