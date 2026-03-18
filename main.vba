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
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                
                ' テキストボックスのみ
                If shp.Type = msoTextBox Then
                    
                    ' 一旦大きくする
                    shp.Width = 1000
                    shp.Height = 1000
                    
                    ' 折り返しON
                    shp.TextFrame2.WordWrap = msoTrue
                    
                    shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                    
                End If
                
            End If

        End If
        
    Next shp

End Sub
