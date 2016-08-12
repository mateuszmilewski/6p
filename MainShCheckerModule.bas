Attribute VB_Name = "MainShCheckerModule"
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
