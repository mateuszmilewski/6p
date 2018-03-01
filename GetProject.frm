VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GetProject 
   Caption         =   "Get content from another project"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8415
   OleObjectBlob   =   "GetProject.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GetProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public newLink As T_Link


Private proj As String
Private plt As String
Private faza As String
Private cw As String


Private Sub BtnSubmit_Click()
    Dim linkedStr As String
    mproj = Trim(CStr(proj))
    mplt = Trim(CStr(plt))
    mfaza = Trim(CStr(faza))
    mcw = Trim(CStr(cw))
    
    Dim sourceLink As T_Link
    Set sourceLink = New T_Link
    
    
    sourceLink.zrob_mnie_z_argsow CStr(mproj), CStr(mplt), CStr(mfaza), CStr(mcw)
    getDataFromThisProjectIntoANewOne newLink, sourceLink
End Sub

Private Sub ListBoxProjects_Click()

    innerChangeStates

End Sub


Private Sub innerChangeStates()


    For X = 0 To Me.ListBoxProjects.ListCount - 1
    
        If Me.ListBoxProjects.Selected(X) Then
            Me.ListBoxPLT.Selected(X) = True
            Me.ListBoxFaza.Selected(X) = True
            Me.ListBoxCW.Selected(X) = True
            
            
            proj = CStr(Me.ListBoxProjects.List(X))
            plt = CStr(Me.ListBoxPLT.List(X))
            faza = CStr(Me.ListBoxFaza.List(X))
            cw = CStr(Me.ListBoxCW.List(X))
            
            
        Else
            Me.ListBoxPLT.Selected(X) = False
            Me.ListBoxFaza.Selected(X) = False
            Me.ListBoxCW.Selected(X) = False
        End If
    Next X
End Sub

Private Sub ListBoxProjects_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    innerChangeStates
End Sub


Private Sub getDataFromThisProjectIntoANewOne(nL As T_Link, sL As T_Link)

    ' w tym subie bede uruchamial kopiowanie danych na bardzo podobnej zasadzie do kopiowania jakie znajduje sie w formularzu nowego projektu uruchamianego z poziomu
    ' przycisku PROJECT MANAGMENT
    ' bedzie tez troche dziwnie, poniewaz suba, ktorego tutaj potrzebuje znajduje sie jego implementacja wewnatrz kodu forma
    ' nie chce tego separowac, wiec tylko upublicznie tego suba i go uruchomie z tamtego poziomu
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    NewProj.dane_dla_nowego_skopiuj_ze_starego sL, nL, ThisWorkbook.Sheets(SIXP.G_main_sh_nm)
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
End Sub
