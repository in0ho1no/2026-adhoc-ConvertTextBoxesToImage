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

Sub DisconnectAllConnectors()

    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        
        If shp.Connector Then
            
            On Error Resume Next
            
            shp.ConnectorFormat.BeginDisconnect
            shp.ConnectorFormat.EndDisconnect
            
            On Error GoTo 0
            
        End If
        
    Next shp

End Sub

Sub ConvertTextBoxToRectangle()

    Dim shp As Shape
    Dim newShp As Shape
    Dim i As Long
    
    Application.ScreenUpdating = False
    
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
            
            With newShp.TextFrame2.TextRange
                
                ' テキスト設定
                .Text = txt
                
                ' フォント設定
                .Font.Name = shp.TextFrame2.TextRange.Font.Name
                .Font.Size = shp.TextFrame2.TextRange.Font.Size
                
                ' 文字色：黒固定
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                
            End With
            
            ' 折り返し
            newShp.TextFrame2.WordWrap = shp.TextFrame2.WordWrap
            
            ' AutoSize
            newShp.TextFrame2.AutoSize = shp.TextFrame2.AutoSize
            
            newShp.Fill.Visible = msoFalse
            
            newShp.Line.Visible = msoFalse
            
            ' 元削除
            shp.Delete
            
        End If
        
    Next i
    
    Application.ScreenUpdating = True

End Sub

Sub OptimizeShapesInSheet()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo Cleanup
    
    ' ① グループ解除
    Call UngroupAllShapesRecursive
    
    ' ② コネクタ切断
    Call DisconnectAllConnectors
    
    ' ③ テキスト調整
    Call ResizeTextBoxesToFitText
    
    ' ④ テキストボックス置換
    Call ConvertTextBoxToRectangle

Cleanup:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
