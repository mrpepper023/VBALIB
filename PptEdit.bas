Attribute VB_Name = "PptEdit"
'test
Private Sub test()

    Debug.Print ActivePresentation.Name

End Sub

'test
Private Sub testselection()

    With ActiveWindow.Selection
        If .Type = ppSelectionNone Then
            Debug.Print "�����I������Ă��܂���"
        End If
        If .Type >= ppSelectionSlides Then
            Debug.Print "�X���C�h���I������Ă��܂�"
        End If
        If .Type >= ppSelectionShapes Then
            Debug.Print "�V�F�C�v���I������Ă��܂�"
            Debug.Print "�I���V�F�C�v���F" & .shaperange.Count
            For Each sset In .shaperange
                Debug.Print "�I���V�F�C�v���F" & sset.Name
                If sset.HasTextFrame Then
                    Debug.Print "hastext"
                    Debug.Print sset.TextFrame.TextRange.text
                End If
                If sset.HasTable Then
                    Debug.Print "hastable"
                    result = InputBox("���䗦����͂��Ă�������")
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
            Debug.Print "�e�L�X�g�͈͂��I������Ă��܂�"
        End If
    End With

End Sub

Sub �e�[�u���Z��������()

    With ActiveWindow.Selection
        If .Type >= ppSelectionShapes Then
            Debug.Print "�V�F�C�v���I������Ă��܂�"
            Debug.Print "�I���V�F�C�v���F" & .shaperange.Count
            For Each sset In .shaperange
                Debug.Print "�I���V�F�C�v���F" & sset.Name
                If sset.HasTable Then
                    Debug.Print "hastable"
                    result = InputBox("���䗦����͂��Ă�������" & vbLf _
                    & "��F1,2,30%,3,F" & vbLf _
                    & " F�����݂̕�����ς��Ȃ�" & vbLf _
                    & " ����%���\�̑S�̕��̐���%" & vbLf _
                    & " �������c�����S�̕��ɑ΂��Đ��l�̔䗦�Ŕz��" & vbLf _
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
                        MsgBox "���z���ƂȂ�񂪂Ȃ��̂ŁA�������ς���Ă��܂��܂�"
                    End If
                    If fixedwidth > entirewidth Then
                        MsgBox "�Œ�v�f�����ŁA���̉����𒴂��Ă��܂��܂�"
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
