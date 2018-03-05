VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SixPCollectionCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' FORREST SOFTWARE
' Copyright (c) 2015 Mateusz Forrest Milewski
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



Private parentPath As String
Private postfixName As String


Private objFSO As Object
Private objFolder As Object
Private objFile As Object
Private i As Integer




Private fileCollection As Collection




Public Function countCollection() As Long
    countCollection = CLng(fileCollection.Count)
End Function

Public Function getCollection() As Collection
    Set getCollection = fileCollection
End Function

Public Sub refreshCollection(c As Collection)
    Set fileCollection = Nothing
    Set fileCollection = c
End Sub



Private Sub Class_Initialize()
    
    parentPath = CStr(SIXP.G_PATH_FOR_SEARCHING)
    postfixName = CStr(SIXP.G_6P_POSTFIX)
    
    Set fileCollection = New Collection
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    SIXP.LoadingFormModule.showLoadingForm

    
End Sub

Private Sub Class_Terminate()
    Set fileCollection = Nothing
    
    SIXP.LoadingFormModule.hideLoadingForm
End Sub




Public Function stworz_kolekcje(my_path As String) As Boolean





    Set fileCollection = New Collection

    Set objFolder = Nothing
    On Error Resume Next
    Set objFolder = objFSO.GetFolder(CStr(my_path))

    If Not objFolder Is Nothing Then
        rekurencyjne_przegladanie_katalogow CStr(my_path), "*" & CStr(postfixName)
    End If
End Function

Private Sub rekurencyjne_przegladanie_katalogow(s As String, pattern As String, Optional oldPattern As String)


    Dim objtmpfolder As Object
    Set objFolder = objFSO.GetFolder(CStr(s))
    
    
    SIXP.LoadingFormModule.incLoadingForm

    For Each objFile In objFolder.Files
    
    
        

        Dim boolCondition As Boolean
        
        If oldPattern <> "" Then
            boolCondition = (CStr(objFile.name) Like CStr(pattern)) Or (CStr(objFile.name) Like CStr(oldPattern))
        Else
            boolCondition = (CStr(objFile.name) Like CStr(pattern))
        End If
        
        If boolCondition Then
            ' Debug.Print CStr(objFile.Path)
            
            jest_juz_taki_i_go_nie_dodamy = False
            
            If fileCollection.Count > 0 Then
                For Each ic In fileCollection
                    If CStr(ic) = CStr(objFile.Path) Then
                        jest_juz_taki_i_go_nie_dodamy = True
                    End If
                Next ic
            End If
            
            If jest_juz_taki_i_go_nie_dodamy = False Then
                fileCollection.Add CStr(objFile.Path)
            End If
            
        End If
        '    'print file path
        '    Cells(i + 1, 2) = objFile.Path
        '    i = i + 1
    Next objFile
    
    If objFolder.SubFolders.Count > 0 Then

        For Each objtmpfolder In objFolder.SubFolders
            rekurencyjne_przegladanie_katalogow CStr(s) & Application.PathSeparator & CStr(objtmpfolder.shortName), pattern, oldPattern
        Next objtmpfolder
    End If

End Sub

