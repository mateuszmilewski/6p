Attribute VB_Name = "FindAll6PModule"
' FORREST SOFTWARE
' Copyright (c) 2018 Mateusz Forrest Milewski
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



Public Sub find_all_6p(ictrl As IRibbonControl)
    
    
    ' main start to look for all files with 6p postfix in fma folder
    ' same logic in collector file
    ' --------------------------------------------------------------------------
    ' --------------------------------------------------------------------------
    Dim c As Collection
    Dim creator As SixPCollectionCreator
    Set creator = New SixPCollectionCreator
    
    With creator
        .stworz_kolekcje SIXP.G_PATH_FOR_SEARCHING
        ' .stworz_kolekcje "C:\WORKSPACE\macros\LESS\6p\"
        Set c = .getCollection()
    End With
    
    Set creator = Nothing
    
    
    'For Each el In c
    '    Debug.Print el
    'Next el
    
    With SIXP.Form6pList
        .ListBoxIn.Clear
        .ListBoxOut.Clear
        
        For Each el In c
            .ListBoxIn.addItem CStr(el)
        Next el
        
        .Show vbModeless
    End With
    
    
    
    ' --------------------------------------------------------------------------
    ' --------------------------------------------------------------------------
End Sub
