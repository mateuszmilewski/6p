VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private s As Shape
Private tf2 As TextFrame2
Private innerTextBox As Variant
Private sh As Worksheet




Public Sub addOpenIssueTextBoxForCell(r As Range, Optional initialText As String)



    Set sh = r.Parent
    
    
    Set s = Nothing
    
    On Error Resume Next
    Set s = sh.Shapes("TB_" & Replace(r.Address, "$", ""))
    
    If s Is Nothing Then
        
        ' tutaj smiala zagryka by textbox jako koment pojawial sie do komorki w jej dowlnym-prawym rogu
        Set s = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            r.Offset(0, 1).Left, r.Offset(0, 1).Top, r.Width * 10, r.Height * 0.9)
            
        s.name = "TB_" & Replace(r.Address, "$", "")
        
        If Trim(initialText) <> "" Then
            Set tf2 = s.TextFrame2
            
            
            tf2.AutoSize = msoAutoSizeShapeToFitText
            tf2.WordWrap = msoFalse
            tf2.TextRange.Characters.Font.name = "Lucida Sans Typewriter"
            
            tf2.TextRange.Characters.Font.Fill.Transparency = 0.4
            tf2.TextRange.Characters.Font.Size = 10
            tf2.TextRange.Characters.Font.Bold = msoCTrue
            
            tf2.TextRange.Text = Trim(initialText)


            
            
            
            
            
            

            
        End If
    Else
        ' w tej sekcji znalezlismy sie poniewaz okazuje sie ze juz taki textbox istnieje
        ' zatem:
        
        Set s = sh.Shapes("TB_" & Replace(r.Address, "$", ""))
        
        tf2.AutoSize = msoAutoSizeShapeToFitText
        tf2.TextRange.Characters.Font.name = "Lucida Sans Typewriter"
        tf2.TextRange.Text = tf2.TextRange.Text & Trim(initialText)
        
    End If
        
    ' to dla innego suba
    ' jednak dalej to mi sie nie podoba!
    's.Select
    
    ' Debug.Print TypeName(Selection) ' OK
    
    'Set tb = Selection
    
    'tb.Formula = "=A1"

    

End Sub


Public Sub addBasicTextBoxForCell(r As Range, Optional initialText As String)



    Set sh = r.Parent
    
    
    Set s = Nothing
    
    On Error Resume Next
    Set s = sh.Shapes("TB_" & Replace(r.Address, "$", ""))
    
    If s Is Nothing Then
        
        ' tutaj smiala zagryka by textbox jako koment pojawial sie do komorki w jej dowlnym-prawym rogu
        Set s = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            r.Offset(0, 1).Left, r.Offset(0, 1).Top, r.Width * 10, r.Height * 0.9)
            
        s.name = "TB_" & Replace(r.Address, "$", "")
        
        If Trim(initialText) <> "" Then
            Set tf2 = s.TextFrame2
            
            tf2.WordWrap = msoFalse
            tf2.TextRange.Characters.Font.name = "Lucida Sans Typewriter"
            
            tf2.TextRange.Characters.Font.Fill.Transparency = 0.4
            tf2.TextRange.Characters.Font.Size = 8
            tf2.TextRange.Characters.Font.Bold = msoCTrue
            
            tf2.TextRange.Text = Trim(initialText)
            
            
            '        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "b"
            '        Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 1).ParagraphFormat. _
            '            FirstLineIndent = 0
            '        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 1).Font
            '            .NameComplexScript = "+mn-cs"
            '            .NameFarEast = "+mn-ea"
            '            .Fill.Visible = msoTrue
            '            .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
            '            .Fill.ForeColor.TintAndShade = 0
            '            .Fill.ForeColor.Brightness = 0
            '            .Fill.Transparency = 0
            
        End If
    Else
        ' w tej sekcji znalezlismy sie poniewaz okazuje sie ze juz taki textbox istnieje
        ' zatem:
        
        Set s = sh.Shapes("TB_" & Replace(r.Address, "$", ""))
        tf2.WordWrap = msoFalse
        tf2.TextRange.Characters.Font.name = "Lucida Sans Typewriter"
        tf2.TextRange.Characters.Font.Size = 8
        
        tf2.TextRange.Text = tf2.TextRange.Text & Trim(initialText)
    End If
        
    ' to dla innego suba
    ' jednak dalej to mi sie nie podoba!
    's.Select
    
    ' Debug.Print TypeName(Selection) ' OK
    
    'Set tb = Selection
    
    'tb.Formula = "=A1"
    'Set tf2 = s.TextFrame2
    'tf2.TextRange.Characters.Text = "=A1"

    

End Sub


Public Sub showTextBoxByName(rRef As Range, nm As String, sh As Worksheet)

    Dim el As Variant
    
    For Each el In sh.Shapes
        
        If el.Type = msoTextBox Then
            If el.name = nm Then
            
                Set s = el
            
                If s.name Like "TB_W*" Then
                    s.Top = rRef.Top + 14
                    s.Left = rRef.Left - 600
                    s.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    s.TextFrame2.WordWrap = True
                End If
                
                If s.name Like "TB_V*" Then
                    s.Top = rRef.Top + 16
                    s.Left = rRef.Left - 600
                    s.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    s.TextFrame2.WordWrap = True
                End If
                
                If s.name Like "TB_S*" Then
                    s.Top = rRef.Top + 16
                    s.Left = rRef.Left
                    s.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    s.TextFrame2.WordWrap = True
                End If
                
                
                If s.name Like "TB_T*" Then
                    s.Top = rRef.Top + 16
                    s.Left = rRef.Left - 100
                    s.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    s.TextFrame2.WordWrap = True
                End If
                
                
                s.Visible = True
                ' Exit Sub
            Else
                Set s = el
                s.Visible = False
            End If
        End If
    Next el
End Sub
