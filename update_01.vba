Option Explicit

' =========================
' メイン処理
' =========================
Sub OptimizeWithGroupPreservation()

    Dim grpInfo As Collection
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    On Error GoTo Cleanup
    
    ' ① グループ構造保存
    Set grpInfo = SaveGroupInfo()
    
    ' ② グループ解除
    Call UngroupAllShapesRecursive
    
    ' ③ テキストボックス調整
    Call ResizeTextBoxesToFitText
    
    ' ④ テキストボックス → 長方形変換（名前維持）
    Call ConvertTextBoxToRectangle
    
    ' ⑤ グループ復元
    Call RestoreGroups(grpInfo)
    
    ' ⑥ グループ単位で画像化
    Call ConvertGroupsToImages

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub


' =========================
' グループ構造保存
' =========================
Function SaveGroupInfo() As Collection

    Dim grpInfo As New Collection
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        
        If shp.Type = msoGroup Then
            
            Dim members As Collection
            Set members = New Collection
            
            Dim i As Long
            For i = 1 To shp.GroupItems.Count
                members.Add shp.GroupItems(i).Name
            Next i
            
            grpInfo.Add members
            
        End If
        
    Next shp
    
    Set SaveGroupInfo = grpInfo

End Function


' =========================
' グループ復元
' =========================
Sub RestoreGroups(grpInfo As Collection)

    Dim members As Collection
    Dim arr() As String
    Dim i As Long
    
    For Each members In grpInfo
        
        ReDim arr(1 To members.Count)
        
        For i = 1 To members.Count
            arr(i) = members(i)
        Next i
        
        On Error Resume Next
        ActiveSheet.Shapes.Range(arr).Group
        On Error GoTo 0
        
    Next members

End Sub


' =========================
' グループ解除（ネスト対応）
' =========================
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


' =========================
' テキストボックスサイズ調整
' =========================
Sub ResizeTextBoxesToFitText()

    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        
        If shp.Type = msoTextBox Then
            If shp.TextFrame2.HasText Then
                
                shp.Width = 500
                shp.Height = 500
                
                shp.TextFrame2.WordWrap = msoTrue
                shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
                
            End If
        End If
        
    Next shp

End Sub


' =========================
' テキストボックス → 長方形変換（名前維持）
' =========================
Sub ConvertTextBoxToRectangle()

    Dim shp As Shape
    Dim newShp As Shape
    Dim i As Long
    
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        
        Set shp = ActiveSheet.Shapes(i)
        
        If shp.Type = msoTextBox Then
            
            Dim l As Double, t As Double, w As Double, h As Double
            l = shp.Left
            t = shp.Top
            w = shp.Width
            h = shp.Height
            
            Dim txt As String
            txt = shp.TextFrame2.TextRange.Text
            
            Dim originalName As String
            originalName = shp.Name
            
            ' 長方形作成
            Set newShp = ActiveSheet.Shapes.AddShape( _
                msoShapeRectangle, l, t, w, h)
            
            ' ★ 名前維持（最重要）
            On Error Resume Next
            newShp.Name = originalName
            On Error GoTo 0
            
            ' テキスト設定
            With newShp.TextFrame2.TextRange
                .Text = txt
                .Font.Name = shp.TextFrame2.TextRange.Font.Name
                .Font.Size = shp.TextFrame2.TextRange.Font.Size
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
            
            newShp.TextFrame2.WordWrap = shp.TextFrame2.WordWrap
            newShp.TextFrame2.AutoSize = shp.TextFrame2.AutoSize
            
            ' 背景透明・枠なし
            newShp.Fill.Visible = msoFalse
            newShp.Line.Visible = msoFalse
            
            shp.Delete
            
        End If
        
    Next i

End Sub


' =========================
' グループを画像化
' =========================
Sub ConvertGroupsToImages()

    Dim shp As Shape
    Dim newShape As Shape
    Dim i As Long
    
    For i = ActiveSheet.Shapes.Count To 1 Step -1
        
        Set shp = ActiveSheet.Shapes(i)
        
        If shp.Type = msoGroup Then
            
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

End Sub