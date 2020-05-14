Attribute VB_Name = "CommentPositionCustomModule"
' https://stackoverflow.com/questions/21465538/how-to-change-the-position-of-a-comment-box-when-hovered-over-it

Option Explicit

Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Dim lngCurPos As POINTAPI
Public CancelHover As Boolean
Dim C4_Left As Double, C4_Right As Double, C4_Top As Double, C4_Bottom As Double

Public Sub superCustomChangeCommentPosition(r As Range)
    CancelHover = False

    With ActiveWindow
        C4_Left = .PointsToScreenPixelsX(r.Left)
        C4_Right = .PointsToScreenPixelsX(r.Offset(0, 1).Left)
        C4_Top = .PointsToScreenPixelsY(r.Top)
        C4_Bottom = .PointsToScreenPixelsY(r.Offset(1, 0).Top)
    End With

    Do
        GetCursorPos lngCurPos

        If lngCurPos.x > C4_Left And lngCurPos.x < C4_Right Then
            If lngCurPos.y > C4_Top And lngCurPos.y < C4_Bottom Then
                '~~> Show the comment forcefully
                r.Comment.Visible = True
                '~~> Re-position the comment. Can use other properties as .Left etc
                r.Comment.Shape.Top = 100
                r.Comment.Shape.Left = 100
            Else
                r.Comment.Visible = False
            End If
        End If

        DoEvents
    Loop Until CancelHover = True
End Sub

