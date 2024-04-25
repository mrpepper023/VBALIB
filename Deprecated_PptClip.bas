Attribute VB_Name = "PptClip"
Function GetTextClip()
    Dim CB As New DataObject
    With CB
        .GetFromClipboard   ''クリップボードからDataObjectにデータを取得する
        GetTextClip = .GetText     ''DataObjectのデータを変数に取得する
    End With
End Function

Function SetTextClip(Optional text = "yokomizo" & vbLf & "michio")
    Dim CB As New DataObject
    With CB
        .SetText text       ''変数のデータをDataObjectに格納する
        .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
    End With
    SetTextClip = text
End Function

'test
Private Sub sample2()
    SetTextClip
End Sub

Function TextTo2DArray(ByRef text)
    splitted = Split(text, vbLf)
    
    colmax = 0
    ubound_splitted = UBound(splitted)
    While splitted(ubound_splitted) = "" Or splitted(ubound_splitted) = vbCr
        ubound_splitted = ubound_splitted - 1
        If ubound_splitted = LBound(splitted) Then
            TextTo2DArray = False
            Exit Function
        End If
    Wend
    
    Dim result() As Variant
    ReDim Preserve result(ubound_splitted, 0)
    Debug.Print LBound(result, 2)

    For i = LBound(splitted) To ubound_splitted
        If Right(splitted(i), 1) = vbCr Then
            splitted(i) = Left(splitted(i), Len(splitted(i)) - 1)
        End If
        temp = Split(splitted(i), vbTab)
        If colmax < UBound(temp) Then
            colmax = UBound(temp)
            ReDim Preserve result(ubound_splitted, colmax)
        End If
        For ii = LBound(temp) To UBound(temp)
            result(i, ii) = temp(ii)
        Next
    Next
    TextTo2DArray = result
End Function


Sub sortbykpi(ByRef dict_key_and_kpi, ByRef sorted_arraylist, Optional descflag = False)
    Set kpi_to_keys = CreateObject("scripting.dictionary")
    Set kpiarray = CreateObject("system.collections.arraylist")
    For Each k In dict_key_and_kpi
        If Not kpi_to_keys.exists(dict_key_and_kpi(k)) Then
            kpi_to_keys.Add dict_key_and_kpi(k), CreateObject("system.collections.arraylist")
            kpiarray.Add dict_key_and_kpi(k)
        End If
        kpi_to_keys(dict_key_and_kpi(k)).Add k
    Next
    kpiarray.sort
    If descflag Then kpiarray.Reverse
    For Each kpi In kpiarray
        For Each origkey In kpi_to_keys(kpi)
            sorted_arraylist.Add origkey
        Next
    Next
    
    Set kpiarray = Nothing
    For Each kpi In kpi_to_keys
        Set kpi_to_keys(kpi) = Nothing
    Next
    Set kpi_to_keys = Nothing
End Sub


Sub 一次元貼付()
    result = TextTo2DArray(GetTextClip())
    '中身表示サンプル
    If IsArray(result) Then
        For r = LBound(result, 1) To UBound(result, 1)
            Debug.Print r & "," & 0 & ": " & result(r, 0)
        Next
    Else
        End
    End If
    '図形リストアップ
    Set selecttable = analize_selection("findtable")
    If selecttable.Count > 0 Then
        '選択範囲は表
        num = LBound(result, 1)
        For Each t In selecttable
            For r = 1 To selecttable(t).Table.Rows.Count
                For c = 1 To selecttable(t).Table.Columns.Count
                    If selecttable(t).Table.Cell(r, c).Selected Then
                        selecttable(t).Table.Cell(r, c).shape.TextFrame.TextRange.text = result(num, 0)
                        num = num + 1
                        If num > UBound(result, 1) Then
                            GoTo skip
                        End If
                    End If
                Next
            Next
        Next
skip:
    Else
        '選択範囲はそれ以外
        Set targetboxes = analize_selection("hastextframe")
        Set dict_name_to_position = CreateObject("scripting.dictionary")
        Set sortedkey = CreateObject("system.collections.arraylist")
        For Each n In targetboxes
            dict_name_to_position.Add n, targetboxes(n).Top + targetboxes(n).Left
        Next
        sortbykpi dict_name_to_position, sortedkey
        
        num = LBound(result, 1)
        For Each n In sortedkey
            targetboxes(n).TextFrame.TextRange.text = result(num, 0)
            num = num + 1
            If num > UBound(result, 1) Then
                Exit For
            End If
        Next
        
        Set dict_name_to_position = Nothing
        Set sortedkey = Nothing
    End If
    Set selecttable = Nothing
End Sub


Sub プロパティ書込()
    Set dict = CreateObject("scripting.dictionary")
    texts = Split(GetTextClip(), vbLf)
    Set minikeys = CreateObject("system.collections.arraylist")
    temp = Split(texts(0), vbTab)
    For Each st In temp
        v = Replace(st, vbCr, "")
        minikeys.Add v
    Next
    For i = 1 To UBound(texts)
        temp = Split(texts(i), vbTab)
        tgtname = ""
        For w = LBound(temp) To UBound(temp)
            v = Replace(temp(w), vbCr, "")
            If minikeys(w) = "Name" Then
                tgtname = v
            End If
        Next
        If Not dict.exists(tgtname) Then
            dict.Add tgtname, CreateObject("scripting.dictionary")
        End If
        Set tgt = dict(tgtname)
        For w = LBound(temp) To UBound(temp)
            v = Replace(temp(w), vbCr, "")
            If IsNumeric(v) Then v = Val(v)
            tgt.Add minikeys(w), v
        Next
        Set tgt = Nothing
    Next
    
    analize_selection_preset "set", dict
    
    Set minikeys = Nothing
    Set dict = Nothing
End Sub

Sub プロパティ読込()
    Set dict = analize_selection("get")
    text = ""
    Set colsd = CreateObject("scripting.dictionary")
    Set colsa = CreateObject("system.collections.arraylist")
    For Each minidict In dict
        For Each k In dict(minidict)
            If Not colsd.exists(k) Then
                colsd.Add k, colsa.Count
                colsa.Add k
            End If
        Next
    Next
    For Each k In colsa
        text = text & k & vbTab
    Next
    text = text & vbCrLf
    For Each minidict In dict
        For Each k In colsa
            If dict(minidict).exists(k) Then
                text = text & dict(minidict)(k) & vbTab
            Else
                text = text & "蟲" & vbTab
            End If
        Next
        text = text & vbCrLf
    Next
    
    Set colsd = Nothing
    Set colda = Nothing
    Set dict = Nothing
    
    SetTextClip text
End Sub



Function GetActiveSlide()

    With ActiveWindow.Selection
        If .Type >= ppSelectionSlides Then
            nowindex = .SlideRange.SlideIndex
        Else
            nowindex = 0
        End If
    End With
    Set GetActiveSlide = ActivePresentation.Slides(nowindex)

End Function

    

Sub エイリアスR_C()
    二次元エイリアス貼付 True
End Sub
Sub エイリアスC_R()
    二次元エイリアス貼付 False
End Sub


Private Sub 二次元エイリアス貼付(order_R_C)
    result = TextTo2DArray(GetTextClip())
    '中身表示サンプル
    If IsArray(result) Then
        For r = LBound(result, 1) To UBound(result, 1)
            For c = LBound(result, 2) To UBound(result, 2)
                Debug.Print r & "," & c & ": " & result(r, c)
            Next
        Next
    Else
        End
    End If
    
    '２系列を取得
    Dim kwdrow(), kwdcol()
    kwdrows = 0
    For r = LBound(result, 1) To UBound(result, 1)
        c = LBound(result, 2)
        If Trim(result(r, c)) <> "" Then kwdrows = kwdrows + 1
    Next
    ReDim Preserve kwdrow(kwdrows - 1)
    i = 0
    For r = LBound(result, 1) To UBound(result, 1)
        c = LBound(result, 2)
        If Trim(result(r, c)) <> "" Then
            kwdrow(i) = Trim(result(r, c))
            Debug.Print kwdrow(i)
            i = i + 1
        End If
    Next
    
    kwdcols = 0
    For r = LBound(result, 1) To UBound(result, 1)
        c = UBound(result, 2)
        If Trim(result(r, c)) <> "" Then kwdcols = kwdcols + 1
    Next
    ReDim Preserve kwdcol(kwdcols - 1)
    i = 0
    For r = LBound(result, 1) To UBound(result, 1)
        c = UBound(result, 2)
        If Trim(result(r, c)) <> "" Then
            kwdcol(i) = Trim(result(r, c))
            Debug.Print kwdcol(i)
            i = i + 1
        End If
    Next
    
    '図形リストアップ
    Set selecttable = analize_selection("findtable")
    Set activenotetextrange = GetActiveSlide().NotesPage.Shapes.Placeholders(2).TextFrame.TextRange

    If selecttable.Count > 0 Then
        '選択範囲は表
        aliasnum = &H100
        num = LBound(result, 1)
        For Each t In selecttable
            sr = 0
            activerow = False
            For r = 1 To selecttable(t).Table.Rows.Count
                sc = 0
                For c = 1 To selecttable(t).Table.Columns.Count
                    If selecttable(t).Table.Cell(r, c).Selected Then
                        If sc < kwdcols And sr < kwdrows Then
                            aliasstr = "|" & Hex(aliasnum)
                            selecttable(t).Table.Cell(r, c).shape.TextFrame.TextRange.text = aliasstr
                            If order_R_C Then
                                activenotetextrange.text = activenotetextrange.text & vbLf & aliasstr & "=" & kwdrow(sr) & "_" & kwdcol(sc)
                            Else
                                activenotetextrange.text = activenotetextrange.text & vbLf & aliasstr & "=" & kwdcol(sc) & "_" & kwdrow(sr)
                            End If
                            
                            aliasnum = aliasnum + 1
                            sc = sc + 1
                            activerow = True
                        End If
                    End If
                Next
                If activerow Then sr = sr + 1
            Next
        Next
    Else
    End If
End Sub
