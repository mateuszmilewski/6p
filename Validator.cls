VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2020 Mateusz Forrest Milewski
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


Public pass As Boolean


Private kolekcjaDanych As Collection
Private vi As ValidationItem



Public pStr_checkIfYYYYCW As String ' = "checkIfYYYYCW"
' to be implemented!!!
Public pStr_checkIfPattern As String
Public pStr_checkIfNumber As String


Public Sub dodajDoKolekcji(el As Variant, callback As String)
    
    Set vi = Nothing
    Set vi = New ValidationItem
    Set vi.frmFld = el
    vi.callback = CStr(callback)
    
    kolekcjaDanych.Add vi
End Sub


Private Sub Class_Initialize()


    pStr_checkIfYYYYCW = "checkIfYYYYCW"
    pStr_checkIfPattern = "checkIfPattern"
    pStr_checkIfNumber = "checkIfNumber"
    
    ' init for validator
    pass = False
    
    Set vi = Nothing
    Set kolekcjaDanych = New Collection
End Sub

Private Sub Class_Terminate()

    ' destruktor validatora
    pass = False
    
    Set kolekcjaDanych = Nothing
    Set vi = Nothing
End Sub



Private Sub checkIfNumber(s As String, frmFld As Variant)

    If IsNumeric(s) Then
        pass = True
    ElseIf Trim(s) = "" Then
        pass = True
    Else
        pass = False
        
        If Not IsMissing(frmFld) Then
            MsgBox frmFld.name & " need to be in format numeric only!!!"
        End If
    End If
    
End Sub

Private Sub checkIfPattern(s As String, ptrn As String, Optional frmFld As Variant)

    

    ' example!
    ' ptrn = "^[0-9]{1,3}"
    
    Dim re As RegExp
    Set re = New RegExp
    
    With re
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = ptrn
    End With
    
    If re.test(s) Then
        pass = True
    Else
        pass = False
        
        If Not IsMissing(frmFld) Then
            MsgBox frmFld.name & " need to be inline with regex pattern!!!"
        End If
        
    End If
    
End Sub


Private Sub checkIfYYYYCW(s As String, Optional frmFld As Variant)

    pass = False

    If Len(s) = 6 Then
        
        If IsNumeric(s) Then
            pass = True
            Exit Sub
        End If
        
        
    ElseIf Trim(s) = "" Then
    
        pass = True
        Exit Sub
    End If
    
    
    
    
    If pass = False Then
        If Not IsMissing(frmFld) Then
            MsgBox frmFld.name & " need to be in format YYYYCW!!!"
        End If
    End If
End Sub




Public Sub run()
    
    ' starting point!
    pass = False
    
    
    For Each vi In kolekcjaDanych
        
        Debug.Print "walidacja dla: " & vi.frmFld.name & " nazwa foo walidujacego: " & vi.callback
        
        If vi.callback = "checkIfPattern" Then
        
            ' to be implemented!!!
            ' Application.run vi.callback, CStr(vi.frmFld.Value), vi.frmFld
            
            
        ElseIf vi.callback = "checkIfYYYYCW" Then
        
            checkIfYYYYCW CStr(vi.frmFld.Value), vi.frmFld
            
            
        ElseIf vi.callback = "checkIfNumber" Then
            
            checkIfNumber CStr(vi.frmFld.Value), vi.frmFld
        
        
        Else
            ' Application.run vi.callback, CStr(vi.frmFld.Value), vi.frmFld
            MsgBox "nie ma takiej mozliwosci - brak zdefiniowanego walidatora!"
        End If
        
        
        If pass = False Then
            Exit Sub
        End If
    Next vi
End Sub


