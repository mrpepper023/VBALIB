Attribute VB_Name = "UtilData"
'=========================================================================================
'UtilData 20230527
'
'UtilDataは主にデータを扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'組み込みのスカラ型とその多次元配列、および、DictionaryとArrayListを文字列にする
'Public Function Serialize(ByRef v) As String
'組み込みのスカラ型とその多次元配列、および、DictionaryとArrayListを文字列から戻す
'Public Function Deserialize(ByRef sstr) As Variant
'指定したデリミタと\a～\n,\\までをエスケープする
'Public Function EscapeDelim(txt, delim)
'指定したデリミタと\a～\n,\\までをエスケープした文字列をデリミタで連結したものを分割する
'Public Function EscapedSplit(txt, delim)
'多次元配列の次元数を得る
'Public Function GetDimension(ByRef ArrayData)
'=========================================================================================


'シリアライズ／デシリアライズ
'エスケープ付きSplit
'エスケープ処理

Private Sub test_ser_deser()

    ReDim arr3d(0 To 3, 0 To 3, 0 To 3) As String
    For i = 0 To 3
    For J = 0 To 3
    For k = 0 To 3
        arr3d(i, J, k) = i & J & k
    Next
    Next
    Next

    Set DT = CreateObject("scripting.dictionary")
    DT.Add 1, arr3d
    Debug.Print DT.Count
    st = Serialize(DT)
    Debug.Print st
    
    '2回呼ぶのはテストの場合のみ
    '2回呼ぶのは実用的でないが、実際には根本は常にDICTIONARYだと思うので、オブジェクト決め打ちで良いはず
    If IsObject(Deserialize(st)) Then
        Set vl = Deserialize(st) 'ふつうはこれで
    Else
        vl = Deserialize(st)
    End If
    
    Stop

End Sub


Public Function Serialize(ByRef v) As String

    result = ""
    
    SerializeRecur result, v
    
    Serialize = Left(result, Len(result) - 1) '最後のデリミタを削除する

End Function

    Private Sub SerializeRecur(ByRef result, ByRef v)
    
        Select Case TypeName(v)
        Case "IRegExp2" '対応するかどうか迷い中(確かにプレコンパイルの正規表現を配列とかにしまっておきたいケースはある）
        
        Case "Dictionary"
            '型名、要素数、（キー、値）×要素数
            result = result + TypeName(v) + "|"
            result = result + CStr(v.Count) + "|"
            For Each k In v
                SerializeRecur result, k
                SerializeRecur result, v.Item(k)
            Next
        Case "ArrayList"
            '型名、要素数、値×要素数
            result = result + TypeName(v) + "|"
            result = result + CStr(v.Count) + "|"
            For Each k In v
                SerializeRecur result, k
            Next
        
       
        Case "String"
            '型名、値
            result = result + TypeName(v) + "|" + EscapeDelim(CStr(v), "|") + "|"
        Case "Date"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Byte"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Currency"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Long"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Integer"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Double"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Single"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Boolean"
            '型名、値
            result = result + TypeName(v) + "|" + CStr(v) + "|"
        Case "Empty"
            '型名
            result = result + TypeName(v) + "|"
        Case "Nothing"
            'Object
            '型名
            result = result + TypeName(v) + "|"
        Case "Null"
            '型名
            result = result + TypeName(v) + "|"
        Case Else
            Debug.Print "未対応の型を発見：" & TypeName(v)
        End Select
        
    End Sub

Public Function Deserialize(ByRef sstr) As Variant

    token = EscapedSplit(sstr, "|")
    ind = 0
    DeserializeRecur obj, token, ind

    If IsObject(obj) Then
        Set Deserialize = obj
    Else
        Deserialize = obj
    End If
    
End Function

    Private Sub DeserializeRecur(ByRef obj, ByRef token, ByRef ind)
    
        'token(ind)を参照しつつobjを構築する
        stype = token(ind)
        ind = ind + 1
        Select Case stype
        Case "IRegExp2" '対応するかどうか迷い中
        
        Case "Dictionary"
            '型名、要素数、（キー、値）×要素数
            Set obj = CreateObject("scripting.dictionary")
            cnt = CLng(token(ind))
            ind = ind + 1
            For i = 1 To cnt
                DeserializeRecur k, token, ind
                DeserializeRecur v, token, ind
                obj.Add k, v
            Next
        
        Case "ArrayList"
            '型名、要素数、値×要素数
            Set obj = CreateObject("system.collections.arraylist")
            cnt = CLng(token(ind))
            ind = ind + 1
            For i = 1 To cnt
                DeserializeRecur k, token, ind
                obj.Add k
            Next
        
           
        Case "String"
            '型名、値
            obj = token(ind)
            ind = ind + 1
        Case "Date"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CDate(v)
        Case "Byte"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CByte(v)
        Case "Currency"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CCur(v)
        Case "Long"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CLng(v)
        Case "Integer"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CInt(v)
        Case "Double"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CDbl(v)
        Case "Single"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CSng(v)
        Case "Boolean"
            '型名、値
            v = token(ind)
            ind = ind + 1
            obj = CBool(v)
        Case "Empty"
            '型名
            obj = Empty
        Case "Nothing"
            'Object
            '型名
            Set obj = Nothing
        Case "Null"
            '型名
            obj = Null
        Case Else
            Debug.Print "未対応の型を発見：" & stype
        End Select
    
    End Sub


Public Function EscapeDelim(txt, delim)

    w = ""
    maxi = Len(txt)
    esc = False
    For i = 1 To maxi
        ch = Mid(txt, i, 1)
        Select Case ch
        Case delim: w = w + "\" + delim
        Case "\": w = w + "\\"
        Case Chr(7): w = w + "\a"
        Case Chr(8): w = w + "\b"
        Case Chr(9): w = w + "\t"
        Case Chr(10): w = w + "\n"
        Case Chr(11): w = w + "\v"
        Case Chr(12): w = w + "\f"
        Case Chr(13): w = w + "\r"
        Case Else
            w = w + ch
        End Select
    Next
    EscapeDelim = w

End Function

' エスケープシーケンス付きのSplit
' \デリミタ, \\(92), \a(7), \b(8), \t(9), \n(10), \v(11), \f(12), \r(13)
Public Function EscapedSplit(txt, delim)
    
    Set temp = CreateObject("system.collections.arraylist")
    w = ""
    maxi = Len(txt)
    esc = False
    For i = 1 To maxi
        ch = Mid(txt, i, 1)
        If esc Then
            Select Case ch
            Case delim: w = w + delim
            Case "\": w = w + "\"
            Case "a": w = w + Chr(7)
            Case "b": w = w + Chr(8)
            Case "t": w = w + Chr(9)
            Case "n": w = w + Chr(10)
            Case "v": w = w + Chr(11)
            Case "f": w = w + Chr(12)
            Case "r": w = w + Chr(13)
            Case Else
                Err.Raise Number:=700, Description:="エラーが発生！"
            End Select
            esc = False
        Else
            Select Case ch
            Case "\": esc = True
            Case delim
                temp.Add w
                w = ""
            Case Else
                w = w + ch
            End Select
        End If
    Next
    temp.Add w

    EscapedSplit = temp.ToArray

End Function

Private Sub test_escapedsplit()

    arr = EscapedSplit("asd,f\nasrg,argsfg\,nb,xfghgsdtr", ",")
    For i = LBound(arr) To UBound(arr)
        Debug.Print i & ": " & arr(i)
    Next

End Sub



