VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T_Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public project As String
Public plt As String
Public faza As String
Public cw As Long



Public Function porownaj_z_drugim_linkiem(obj As T_Link) As Boolean
    
    porownaj_z_drugim_linkiem = False

    If Me.project = obj.project Then
        If Me.plt = obj.plt Then
            If Me.faza = obj.faza Then
                If Me.cw = obj.cw Then
                    porownaj_z_drugim_linkiem = True
                End If
            End If
        End If
    End If
End Function


Public Function znajdz_siebie_w_arkuszu(sh As Worksheet) As Range
    Set znajdz_siebie_w_arkuszu = Nothing
    
    Dim r As Range
    Set r = sh.Range("A2")
    
    Do
        If CStr(r) = CStr(project) Then
            If CStr(r.Offset(0, 1)) = CStr(plt) Then
                If CStr(r.Offset(0, 2)) = CStr(faza) Then
                    If CLng(r.Offset(0, 3)) = CLng(cw) Then
                        Set znajdz_siebie_w_arkuszu = r
                        Exit Function
                    End If
                End If
            End If
        End If
        Set r = r.Offset(1, 0)
    Loop Until Trim(r) = ""
End Function

Public Sub zrob_mnie_z_arkusza(r As Range)
    
    If r.Column = 1 Then
        project = r
        plt = r.Offset(0, 1)
        faza = r.Offset(0, 2)
        On Error Resume Next
        cw = CLng(r.Offset(0, 3))
    End If
End Sub

' super wazna czesc tej implementacji!
' ================================================================================================
' ================================================================================================
Public Function znajdz_moje_najwczesniejsze_pod_wzdledem_cw_wystapienie(sh As Worksheet) As Range
    Set znajdz_moje_najwczesniejsze_pod_wzdledem_cw_wystapienie = Nothing
End Function

Public Function znajdz_moje_najpozniejsze_pod_wzdledem_cw_wystapienie(sh As Worksheet) As Range
    Set znajdz_moje_najpozniejsze_pod_wzgledem_cw_wystapienie = Nothing
End Function
' ================================================================================================
' ================================================================================================
