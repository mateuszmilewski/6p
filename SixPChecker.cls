VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SixPChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sh As Worksheet
Private r As Range

Public Function sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main() As Boolean

    sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main = False
    
    If Me.sprawdz_czy_arkusz_aktywny_to_ten_arkusz() Then

        Set sh = ActiveSheet
        Set r = sh.Range("A1")
        
        If sprawdz_pierwsze_cztery_kolumny_proj_plt_faza_cw_z_danymi_w_rejestrze(r) Then
            sprawdz_czy_aktywny_arkusz_jest_w_stanie_otworzyc_formularz_form_main = True
        End If
        
        
    End If
End Function


Public Function sprawdz_czy_arkusz_aktywny_to_ten_arkusz() As Boolean
    sprawdz_czy_arkusz_aktywny_to_ten_arkusz = False
    
    
    If ThisWorkbook.FullName = ActiveWorkbook.FullName Then
        sprawdz_czy_arkusz_aktywny_to_ten_arkusz = True
    End If
End Function

Private Function sprawdz_pierwsze_cztery_kolumny_proj_plt_faza_cw_z_danymi_w_rejestrze(ByRef ir As Range) As Boolean
    sprawdz_pierwsze_cztery_kolumny_proj_plt_faza_cw_z_danymi_w_rejestrze = False
    
    Dim reg As Worksheet
    Set reg = ThisWorkbook.Sheets(SIXP.G_register_sh_nm)
    
    
    
    If CStr(reg.Range("B2")) = CStr(ir) Then
        If CStr(reg.Range("B3")) = CStr(ir.Offset(0, 1)) Then
            If CStr(reg.Range("B4")) = CStr(ir.Offset(0, 2)) Then
                If CStr(reg.Range("B5")) = CStr(ir.Offset(0, 3)) Then
                    sprawdz_pierwsze_cztery_kolumny_proj_plt_faza_cw_z_danymi_w_rejestrze = True
                End If
            End If
        End If
    End If
End Function
