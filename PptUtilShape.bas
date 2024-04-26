Attribute VB_Name = "PptUtilShape"
'=========================================================================================
'PptUtilShape 20230527
'
'PptUtilShapeは主にシェイプ操作を扱う、Powerpoint VBA独特のコードを集めたもの
'=========================================================================================
'選択したシェイプのブーリアン演算のラッパー
'Function SelectionMergeShapes(mergecmd As MsoMergeCmd, Optional PrimaryShape As Shape)
'icon2pictureで用いる枠消し用ドーナッツ型の生成
'Sub mydonuts()
'グループ化されたアイコン画像を大きなシェイプに統合
'Sub icon2picture()
'=========================================================================================

'外形のLeft , Top, Width, Height
' -5#, -5#, 550.17, 550.17
'
'穴のLeft , Top, Width, Height
' 23.48, 23.48, 489.91, 489.91
'
'アイコンを拡大する際のLeft , Top, Width, Height
' 0#, 0#, 540#, 540#


Function SelectionMergeShapes(mergecmd As MsoMergeCmd, Optional PrimaryShape As Shape)
    
    ' マージ結果のシェイプを得られない謎仕様のため、予め名前をリストアップしてから、マージする
    Set names = CreateObject("scripting.dictionary")
    For Each sh In ActiveWindow.Selection.SlideRange.Shapes
        names.Add sh.Name, 1
    Next
    
    ' マージする
    ActiveWindow.Selection.ShapeRange.MergeShapes msoMergeSubtract, PrimaryShape

    ' マージ結果のシェイプを得られない謎仕様のため、名前をチェックして、増えているものがあればそれが新しいもの
    For Each sh In ActiveWindow.Selection.SlideRange.Shapes
        If Not names.exists(sh.Name) Then Set SelectionMergeShapes = sh
    Next
    Set names = Nothing
    
End Function

Sub mydonuts()

    Set circle1 = Application.ActivePresentation.Slides(2).Shapes.AddShape(msoShapeOval, -5, -5, 550.17, 550.17)
    Set circle2 = Application.ActivePresentation.Slides(2).Shapes.AddShape(msoShapeOval, 23.48, 23.48, 489.91, 489.91)
    ' 基本的には選択したシェイプをマージすることしかできない謎仕様のため、順序に注意して選択する
    circle1.Select Replace:=msoTrue
    circle2.Select Replace:=msoFalse
    
    ' Wrapperを呼び出す
    Set result = SelectionMergeShapes(msoMergeCombine)
    
    result.Select
    
End Sub


Sub icon2picture()

    If Application.ActivePresentation.Slides(2).Shapes.Count > 0 Then
        MsgBox "作業用スライドが空ではありません"
        End
    End If
    If Application.ActivePresentation.Slides(3).Shapes.Count <> 1 Then
        MsgBox "マスク画像スライドが異常な状態です"
        End
    End If
    For Each sh In Application.ActivePresentation.Slides(3).Shapes
        If sh.Name <> "RoundMask" Then
            MsgBox "マスク画像スライドが異常な状態です"
            End
        End If
    Next

    ActiveWindow.View.GotoSlide Index:=2
    Application.ActivePresentation.Slides(2).Shapes.Paste.Select
    'スライドのサイズを取得
    sld_w = Application.ActivePresentation.Slides(3).Master.Width
    sld_h = Application.ActivePresentation.Slides(3).Master.Height

    '図形のサイズを取得
    With ActiveWindow.Selection.ShapeRange
        Debug.Print .Width
        Debug.Print .Height
        .Width = sld_h
        .Height = sld_h
        .Left = 0
        .Top = 0
        .Name = "Icon3Picture"
        .Ungroup
    End With
    Application.ActivePresentation.Slides(2).Shapes.SelectAll
    ActiveWindow.Selection.ShapeRange.MergeShapes msoMergeUnion
     For Each sh In Application.ActivePresentation.Slides(3).Shapes
        If sh.Name = "RoundMask" Then
            sh.Copy
        End If
    Next
    ActiveWindow.View.GotoSlide Index:=2
    Application.ActivePresentation.Slides(2).Shapes.Paste
    For Each sh In Application.ActivePresentation.Slides(2).Shapes
        If sh.Name <> "RoundMask" Then
            sh.Select
        End If
    Next
    For Each sh In Application.ActivePresentation.Slides(2).Shapes
        If sh.Name = "RoundMask" Then
            sh.Select Replace:=msoFalse
        End If
    Next
    ActiveWindow.Selection.ShapeRange.MergeShapes msoMergeSubtract
    
End Sub
