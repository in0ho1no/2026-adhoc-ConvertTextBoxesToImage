Option Explicit

' =========================
' メイン処理
' =========================
Sub OptimizeByGroup()

    Dim shp As Shape
    Dim i As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo Cleanup
    
    ' 後ろから（削除安全）
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        
        Set shp = ActiveSheet.Shapes(i)
        
        If shp.Type = msoGroup Then
            ProcessSingleGroup shp
        End If
        
    Next i

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


' =========================
' グループ単位処理（安全版）
' =========================
Sub ProcessSingleGroup(grp As Shape)

    Dim l As Double, t As Double
    l = grp.Left
    t = grp.Top
    
    ' グループ解除
    Dim items As ShapeRange
    Set items = grp.Ungroup
    
    Dim shp As Shape
    Dim i As Long
    
    ' ★ Ungroup直後のShapeRangeをそのまま使う
    For i = 1 To items.Count
        
        Set shp = items(i)
        
        If shp.Type = msoTextBox Then
            
            Dim newShp As Shape
            Set newShp = ReplaceTextBox(shp)
            
            ' ★ ShapeRangeを差し替え
            items(i).Name = newShp.Name
            
        End If
        
    Next i
    
    ' ★ 元のShapeRangeだけで再グループ化（混入防止）
    Dim newGrp As Shape
    Set newGrp = items.Group
    
    newGrp.Left = l
    newGrp.Top = t
    
    ' 画像化
    ConvertShapeToImage newGrp

End Sub


' =========================
' テキストボックス置換
' =========================
Function ReplaceTextBox(shp As Shape) As Shape

    Dim l As Double, t As Double, w As Double, h As Double
    l = shp.Left: t = shp.Top
    w = shp.Width: h = shp.Height
    
    Dim txt As String
    txt = shp.TextFrame2.TextRange.Text
    
    Dim newShp As Shape
    Set newShp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, l, t, w, h)
    
    With newShp.TextFrame2.TextRange
        .Text = txt
        .Font.Name = shp.TextFrame2.TextRange.Font.Name
        .Font.Size = shp.TextFrame2.TextRange.Font.Size
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    newShp.TextFrame2.WordWrap = shp.TextFrame2.WordWrap
    newShp.TextFrame2.AutoSize = shp.TextFrame2.AutoSize
    
    newShp.Fill.Visible = msoFalse
    newShp.Line.Visible = msoFalse
    
    shp.Delete
    
    Set ReplaceTextBox = newShp

End Function


' =========================
' 図形を画像化
' =========================
Sub ConvertShapeToImage(shp As Shape)

    Dim l As Double, t As Double, w As Double, h As Double
    l = shp.Left: t = shp.Top
    w = shp.Width: h = shp.Height
    
    shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    DoEvents
    
    ActiveSheet.Paste
    
    Dim newShape As Shape
    Set newShape = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)
    
    With newShape
        .Left = l
        .Top = t
        .Width = w
        .Height = h
    End With
    
    shp.Delete

End Sub