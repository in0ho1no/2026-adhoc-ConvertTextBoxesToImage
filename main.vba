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
        
        ' テキストを持つオブジェクトのみ対象
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                
                ' ★ テキストボックスのみ対象
                If shp.Type = msoTextBox Then
                    
                    With shp.TextFrame
                        
                        ' 一旦オートサイズをOFF
                        .AutoSize = False
                        
                        ' ★ 十分大きくする（上限は適宜調整）
                        shp.Width = 1000
                        shp.Height = 1000
                        
                        ' 折り返しON（重要）
                        .WordWrap = True
                        
                        ' ★ 文字に合わせて自動調整
                        .AutoSize = True
                        
                    End With
                    
                End If
                
            End If
        End If
        
    Next shp

End Sub
