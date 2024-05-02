Attribute VB_Name = "PptUtilEdit"
'test
Private Sub test()

    Debug.Print ActivePresentation.Name

End Sub

'test
Private Sub testselection()

    With ActiveWindow.Selection
        If .Type = ppSelectionNone Then
            Debug.Print "何も選択されていません"
        End If
        If .Type >= ppSelectionSlides Then
            Debug.Print "スライドが選択されています"
        End If
        If .Type >= ppSelectionShapes Then
            Debug.Print "シェイプが選択されています"
            Debug.Print "選択シェイプ数：" & .shaperange.Count
            For Each sset In .shaperange
                Debug.Print "選択シェイプ名：" & sset.Name
                If sset.HasTextFrame Then
                    Debug.Print "hastext"
                    Debug.Print sset.TextFrame.TextRange.text
                End If
                If sset.HasTable Then
                    Debug.Print "hastable"
                    result = InputBox("幅比率を入力してください")
                    With sset.Table
                    End With
                End If
                If sset.HasSmartArt Then
                    Debug.Print "hassmartart"
                End If
                If sset.HasChart Then
                    Debug.Print "haschart"
                End If
            Next
        End If
        If .Type >= ppSelectionText Then
            Debug.Print "テキスト範囲が選択されています"
        End If
    End With

End Sub

Sub テーブルセル幅調整()

    With ActiveWindow.Selection
        If .Type >= ppSelectionShapes Then
            Debug.Print "シェイプが選択されています"
            Debug.Print "選択シェイプ数：" & .shaperange.Count
            For Each sset In .shaperange
                Debug.Print "選択シェイプ名：" & sset.Name
                If sset.HasTable Then
                    Debug.Print "hastable"
                    result = InputBox("幅比率を入力してください" & vbLf _
                    & "例：1,2,30%,3,F" & vbLf _
                    & " F→現在の幅から変えない" & vbLf _
                    & " 数字%→表の全体幅の数字%" & vbLf _
                    & " 数字→残った全体幅に対して数値の比率で配分" & vbLf _
                    )
                    If result = False Or Len(Trim(result)) = 0 Then End
                    temp = Split(result, ",")
                    If UBound(temp) - LBound(temp) + 1 < sset.Table.Columns.Count Then
                        suffix = Right(result, Len(result) - InStrRev(result, ",") + 1)
                        Debug.Print suffix
                        For i = 1 To sset.Table.Columns.Count - (UBound(temp) - LBound(temp) + 1)
                            result = result + suffix
                        Next
                    End If
                    Debug.Print result
                    result = Split(result, ",")
                    resultsum = 0
                    entirewidth = 0#
                    fixedwidth = 0#
                    min_result_maxc = UBound(result)
                    If min_result_maxc > sset.Table.Columns.Count - 1 Then
                        min_result_maxc = sset.Table.Columns.Count - 1
                    End If
                    For i = LBound(result) To min_result_maxc
                        entirewidth = entirewidth + sset.Table.Columns(i + 1).Width
                    Next
                    For i = LBound(result) To min_result_maxc
                        If UCase(Trim(result(i))) = "F" Then
                            result(i) = "F"
                            fixedwidth = fixedwidth + sset.Table.Columns(i + 1).Width
                        ElseIf Right(Trim(result(i)), 1) = "%" Then
                            percent = Left(Trim(result(i)), Len(Trim(result(i))) - 1)
                            result(i) = percent & "%"
                            fixedwidth = fixedwidth + (entirewidth * percent / 100)
                        Else
                            result(i) = Val(Trim(result(i)))
                            resultsum = resultsum + result(i)
                        End If
                    Next
                    
                    If resultsum = 0 Then
                        MsgBox "比例配分となる列がないので、横幅が変わってしまいます"
                    End If
                    If fixedwidth > entirewidth Then
                        MsgBox "固定要素だけで、元の横幅を超えてしまいます"
                    End If
                    
                    With sset.Table
                        maxc = .Columns.Count
                        For c = 1 To maxc
                            rule = result(c - 1)
                            If rule <> "F" Then
                                If Right(rule, 1) = "%" Then
                                    percent = Val(Left(rule, Len(rule) - 1))
                                    sset.Table.Columns(c).Width = entirewidth * percent / 100
                                ElseIf entirewidth > fixedwidth Then
                                    sset.Table.Columns(c).Width = (entirewidth - fixedwidth) * rule / resultsum
                                End If
                            End If
                        Next
                    End With
                End If
            Next
        End If
    End With

End Sub
