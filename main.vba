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

Sub ConvertTextBoxToRectangle()

    Dim shp As Shape
    Dim newShp As Shape
    Dim i As Long
    
    Application.ScreenUpdating = False
    
    ' 後ろから回す（削除するため）
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        
        Set shp = ActiveSheet.Shapes(i)
        
        If shp.Type = msoTextBox Then
            
            ' 位置・サイズ保持
            Dim l As Double, t As Double, w As Double, h As Double
            l = shp.Left
            t = shp.Top
            w = shp.Width
            h = shp.Height
            
            ' テキスト取得
            Dim txt As String
            txt = shp.TextFrame2.TextRange.Text
            
            ' 長方形作成
            Set newShp = ActiveSheet.Shapes.AddShape( _
                msoShapeRectangle, l, t, w, h)
            
            ' テキスト設定
            newShp.TextFrame2.TextRange.Text = txt
            
            ' 書式コピー（重要）
            newShp.TextFrame2.TextRange.Font.Name = shp.TextFrame2.TextRange.Font.Name
            newShp.TextFrame2.TextRange.Font.Size = shp.TextFrame2.TextRange.Font.Size
            
            ' 折り返し
            newShp.TextFrame2.WordWrap = shp.TextFrame2.WordWrap
            
            ' AutoSizeも引き継ぐ
            newShp.TextFrame2.AutoSize = shp.TextFrame2.AutoSize
            
            ' 塗りつぶし・線をコピー
            newShp.Fill.ForeColor.RGB = shp.Fill.ForeColor.RGB
            newShp.Line.ForeColor.RGB = shp.Line.ForeColor.RGB
            
            ' 元削除
            shp.Delete
            
        End If
        
    Next i
    
    Application.ScreenUpdating = True

End Sub
