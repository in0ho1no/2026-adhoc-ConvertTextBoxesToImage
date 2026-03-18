Sub UngroupAllShapesRecursive()

    Dim hasGroup As Boolean
    Dim shp As Shape
    Dim i As Long

    Do
        hasGroup = False
        
        For i = ActiveSheet.Shapes.Count To 1 Step -1
            Set shp = ActiveSheet.Shapes(i)
            
            If shp.Type = msoGroup Then
                shp.Ungroup
                hasGroup = True
            End If
        Next i
        
    Loop While hasGroup

End Sub

Sub ResizeTextBoxesToFitText()

    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        
        ' テキストあり
        If shp.Type = msoTextBox Then
            If shp.TextFrame2.HasText Then
                
                shp.Width = 500
                shp.Height = 1000
                
                shp.TextFrame2.WordWrap = msoTrue
                shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                
            End If
        End If
        
    Next shp

End Sub

Sub ConvertShapesToImages_Stable()

    Dim shp As Shape
    Dim newShape As Shape
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        
        Set shp = ActiveSheet.Shapes(i)
        
        If shp.Type <> msoPicture Then
            
            Dim l As Double, t As Double, w As Double, h As Double
            l = shp.Left
            t = shp.Top
            w = shp.Width
            h = shp.Height
            
            shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            
            DoEvents
            
            ActiveSheet.Paste
            Set newShape = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)
            
            With newShape
                .Left = l
                .Top = t
                .Width = w
                .Height = h
            End With
            
            shp.Delete
            
        End If
        
    Next i
    
    Application.ScreenUpdating = True

End Sub
