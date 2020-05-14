VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableInTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cls As Long
Private rws As Long


Private labels As Collection
Private sizes As Collection
Public collOfDataLines As Collection
Private line As TableLine
Private ti As TableItem



Public Sub setCols(c As Long)
    cls = CLng(c)
End Sub

Public Sub setRows(r As Long)
    rws = CLng(r)
End Sub


Public Sub defineLables(ParamArray arr() As Variant)
    
    For x = LBound(arr) To UBound(arr)
        
        labels.Add CStr(arr(x))
    Next x
End Sub

Public Sub defineSizes(ParamArray arr() As Variant)
    
    For x = LBound(arr) To UBound(arr)
        
        sizes.Add Int(arr(x))
    Next x
End Sub


Public Sub tryToAddLineOfTheData(ParamArray arr() As Variant)
    
    If Not sizes Is Nothing Then
        
        If sizes.Count > 0 Then
        
            If sizes.Count = labels.Count Then
            
                ' main body of the logic of this sub
                ' ---------------------------------------------------
                
                Dim arrSize As Integer
                arrSize = UBound(arr) - LBound(arr) + 1
                
                If sizes.Count = arrSize Then
                
                
                    ' now really the logic
                    ' -------------------------------------------------------------------
                    ' -------------------------------------------------------------------
                    
                    
                    If Not collOfDataLines Is Nothing Then
                    
                    
                        Set line = Nothing
                        Set line = New TableLine
                        
                        Dim iter As Integer, i As Integer
                        iter = LBound(arr)
                        i = 1
                        
                        For Each l In labels
                        
                            ' ///////////////////////////////////
                            line.addItem restring(CStr(arr(iter)), sizes.item(i), Int(i)), Left(l, 1)
                            ' ///////////////////////////////////
                            i = i + 1
                            iter = iter + 1
                        Next l
                        
                        collOfDataLines.Add line
                    End If
                    
                    
                    ' -------------------------------------------------------------------
                    ' -------------------------------------------------------------------
                
                Else
                    MsgBox "Provided data are not matching size with table!", vbCritical
                    'End
                End If
                
                
                ' ---------------------------------------------------
            Else
                MsgBox "size and labels wrongly defined!", vbCritical
                'End
            End If
        Else
            MsgBox "You defined table with size = 0!", vbCritical
            'End
        End If
    Else
        MsgBox "Size of the table is not defined!", vbCritical
        'End
    End If
End Sub

Private Function restring(txt As String, s As Integer, ktoraKolumna As Integer) As String
    restring = ""
    
    ' s is a limit for the line
    If Len(txt) > s Then
        
        ' restring = Left(txt, s - 1) & "*"
        txtArr = Split(txt, " ")
        
        tmpForLine = ""
        For x2 = LBound(txtArr) To UBound(txtArr)
            
            If Len(tmpForLine & txtArr(x2)) & " " < s Then
                tmpForLine = tmpForLine & txtArr(x2) & " "
                
            ElseIf Trim(tmpForLine) = "" And (txtArr(x2) & " ") > s Then
            
                tmpForLine = tmpForLine & Left((txtArr(x2) & " "), s - 3) & "..."
                
            Else
                ' missing spaces counter
                msCnt = s - Len(tmpForLine)
                missingSpaces = Application.WorksheetFunction.Rept(" ", CDbl(msCnt))
                restring = restring & tmpForLine & missingSpaces & chr(10)
                
                ' Debug.Print Len(tmpForLine & missingSpaces)
                
                tmpForLine = ""
                ' musimy wrocic do tego co zostal odrzucony!
                x2 = x2 - 1
            End If
        Next x2
        
        If tmpForLine <> "" Then
            ' missing spaces counter
            msCnt = s - Len(tmpForLine)
            missingSpaces = Application.WorksheetFunction.Rept(" ", CDbl(msCnt))
            restring = restring & tmpForLine & missingSpaces & chr(10)
            
            ' Debug.Print Len(tmpForLine & missingSpaces)
            
            tmpForLine = ""
        End If

        
    ElseIf Len(txt) = s Then
        restring = txt
    Else
        ' len(txt) < s ' so spaces on left and right
        Dim howManySpaces As Integer
        howManySpaces = s - Int(Len(txt))
        
        onLeft = Int(howManySpaces * 0.5)
        onRight = howManySpaces - onLeft
        
        ' !!!
        '/ ///////////////////////////////////////////////////////////////////////
        restring = Application.WorksheetFunction.Rept(" ", CDbl(onLeft)) & _
            CStr(txt) & Application.WorksheetFunction.Rept(" ", CDbl(onRight))
        '/ ///////////////////////////////////////////////////////////////////////
    End If
End Function



Public Function drawThisTable() As String


    drawThisTable = ""

    ' check sizes and labels
    If sizes.Count = labels.Count Then
        ' simple scenario!
        ' ------------------------------------------
        Dim header As Boolean
        header = True
        headerBreakLine = ""
        For Each line In collOfDataLines
            strLine = ""
            
            
            
            For l2 = 1 To line.maxLines
                For Each k In line.d.Keys
                    ' Debug.Print "Key: " & Key & " Value: " & dic(Key)
                    ' old
                    'takeOnlyFirstOne = Split(line.d(k), Chr(10))(0)
                    'strLine = strLine & " | " & takeOnlyFirstOne
                    
                    Set ti = line.d(k)
                    takeOneItem = ""
                    On Error Resume Next
                    takeOneItem = ti.arr.item(l2)
                    
                    If takeOneItem = "" And l2 > 1 Then
                        takeOneItem = Application.WorksheetFunction.Rept(" ", CDbl(Len(ti.arr.item(1))))
                    End If
                    
                    strLine = strLine & " | " & takeOneItem
                    
                    
                Next
                
                strLine = strLine & chr(10)
            Next l2
            
            ' strLine = strLine & Chr(10)
            If header Then
                headerBreakLine = Application.WorksheetFunction.Rept("-", CDbl(Len(strLine)))
                strLine = strLine & headerBreakLine & chr(10)
                header = False
            End If
            drawThisTable = drawThisTable & strLine & headerBreakLine & chr(10)
        Next line
        
        
        ' drawThisTable
        
        ' ------------------------------------------
    Else
        MsgBox "Sth went wrong during data filling!", vbCritical
    End If

End Function

Private Sub Class_Initialize()
    cls = 0
    rws = 0
    
    Set labels = New Collection
    Set sizes = New Collection
    Set collOfDataLines = New Collection
End Sub

Private Sub Class_Terminate()
    cls = 0
    rws = 0
    Set labels = Nothing
    Set sizes = Nothing
    Set collOfDataLines = Nothing
End Sub
