VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormGetOneItem 
   Caption         =   "Get one item"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3690
   OleObjectBlob   =   "FormGetOneItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormGetOneItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ors section
Public ors_mrd As String
Public ors_build As String
Public ors_bom_freeze As String
Public ors_no_of_veh As String
Public ors_orders_due As String
Public ors_released As String
Public ors_weeks_delay As String


' rbpc section
Public rbpc_num_of_veh As String
Public rbpc_tbw As String
Public rbpc_order_release_changes As String
Public rbpc_comment As String




Private Sub BtnSubmit_Click()
    G_ONE_ITEM_LOGIC_WAITING_FOR_SELECTION_CHANGE = False
    Hide
    
    If Me.LabelClient.Caption = FormOrderReleaseStatus.name Then
    
    
        If FormOrderReleaseStatus.Visible Then
        
        
            With FormOrderReleaseStatus
                .TextBoxBOMFreeze.Value = Me.ors_bom_freeze
                .TextBoxBuild.Value = Me.ors_build
                .TextBoxMRD.Value = Me.ors_mrd
                .TextBoxNoOfVeh.Value = Me.ors_no_of_veh
                .TextBoxOrdersDue.Value = Me.ors_orders_due
                .TextBoxReleased.Value = Me.ors_released
                .TextBoxWeeksDelay.Value = Me.ors_weeks_delay
            End With
        Else
            MsgBox "Form ORS not visible! Sth went really wrong!"
        End If
        
    ElseIf Me.LabelClient.Caption = FormRecentBuildPlanChanges.name Then
    
        If FormRecentBuildPlanChanges.Visible Then
        
            With FormRecentBuildPlanChanges
                .TextBoxCmnt.Value = Me.rbpc_comment
                .TextBoxNoOfVeh.Value = Me.rbpc_num_of_veh
                .TextBoxReleased.Value = Me.rbpc_order_release_changes
                .TextBoxTBW.Value = Me.rbpc_tbw
            End With
        Else
            MsgBox "Form RBPC not visible! Sth went really wrong!"
        End If
    
    End If
End Sub


Private Sub UserForm_Deactivate()
    G_ONE_ITEM_LOGIC_WAITING_FOR_SELECTION_CHANGE = False
End Sub

Private Sub UserForm_Terminate()
    G_ONE_ITEM_LOGIC_WAITING_FOR_SELECTION_CHANGE = False
End Sub


Public Sub clear_rbpc()
    
    With Me
        .rbpc_comment = ""
        .rbpc_num_of_veh = ""
        .rbpc_order_release_changes = ""
        .rbpc_tbw = ""
    End With
End Sub


Public Sub clear_ors()
    
    With Me
        .ors_bom_freeze = ""
        .ors_build = ""
        .ors_mrd = ""
        .ors_no_of_veh = ""
        .ors_orders_due = ""
        .ors_released = ""
        .ors_weeks_delay = ""
    End With
End Sub
