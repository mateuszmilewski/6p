Attribute VB_Name = "MainShCheckerModule"
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


Public Sub sprawdz_main_sh_na_kolumnach_updateow_kolejnych_arkuszy_zrodlowych()
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    
    Dim r As Range, l As T_Link
    Set r = sh.Range("A2")
    
    
    
    Do
        Set l = New T_Link
        l.zrob_mnie_z_range r
        ' teraz metody w l - znajdz moje najpozniejsze i najwczesniejsze wystapienia w arkuszu
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Sub
