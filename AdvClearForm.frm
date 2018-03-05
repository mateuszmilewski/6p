VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdvClearForm 
   Caption         =   "AdvClearForm"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3855
   OleObjectBlob   =   "AdvClearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AdvClearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnAll_Click()


    gotoThisWorkbookMainA1
            
    ans = MsgBox("Are you sure?", vbYesNo)
    
    
    
    If ans = vbYes Then
    
        Hide
        
        SIXP.clear_all_items
        MsgBox "Done!"
    Else
        MsgBox "nothing to do..."
        
    End If
End Sub

Private Sub BtnClearByWildcard_Click()
    ans = MsgBox("Are you sure?", vbYesNo)
    
    If ans = vbYes Then
        
        If Trim(Me.TextBoxWildcard.Value) = "" Then
            ans = MsgBox("puste pole text box spowoduje usuniecie wszystkiego, kontynuowac?", vbYesNo)
            
            If ans = vbYes Then
                SIXP.clear_all_items
                MsgBox "Done!"
            Else
                MsgBox "nothing to do..."
            End If
        Else
            SIXP.clear_by_wildcard CStr(Me.TextBoxWildcard.Value)
            MsgBox "Done!"
        End If
    Else
        MsgBox "nothing to do..."
        
    End If
End Sub

Private Sub BtnClose_Click()
    Hide
End Sub
