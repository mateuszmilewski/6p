Attribute VB_Name = "RemoverModule"
Public Sub removeEntireRow()

    Dim r As Range
    
    Set r = ThisWorkbook.Sheets(SIXP.G_NEW_ONE_PAGER_SH_NM).Range("A40:A500000")
    r.EntireRow.Delete xlupShift
End Sub



