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
        
        '配列のシリアライズはすべて同じ形
        '型名、次元数（要素数列挙）、値×Π要素数列挙（SerializeRecurWithLoop）
        Case "Variant()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Object()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "String()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Byte()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Long()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Integer()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Double()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Single()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Boolean()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Date()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        Case "Currency()"
            result = result + TypeName(v) + "|"
            result = result + MultiDimArray2String(v) + "|"
            SerializeRecurWithLoop result, v
        
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
        
        '配列のデシリアライズはすべて同じ形
        '型名、次元数（要素数列挙）、値×Π要素数列挙（DeserializeRecurWithLoop）
        Case "Variant()"
            obj = String2MultiDimArray_Variant(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Object()"
            obj = String2MultiDimArray_Object(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "String()"
            obj = String2MultiDimArray_String(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Byte()"
            obj = String2MultiDimArray_Byte(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Long()"
            obj = String2MultiDimArray_Long(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Integer()"
            obj = String2MultiDimArray_Integer(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Double()"
            obj = String2MultiDimArray_Double(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Single()"
            obj = String2MultiDimArray_Single(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Boolean()"
            obj = String2MultiDimArray_Boolean(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Date()"
            obj = String2MultiDimArray_Date(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
        Case "Currency()"
            obj = String2MultiDimArray_Currency(token(ind))
            ind = ind + 1
            DeserializeRecurWithLoop obj, token, ind
            
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


'serializerで対応すべきデータ型
Private Sub test_typename()

    'ColorはLong
    reg = RGB(255, 255, 255)
    Debug.Print TypeName(reg)

    'IRegExp2
    Set reg = CreateObject("VBScript.RegExp")
    Debug.Print TypeName(reg)

    'Dictionary
    Set tes = CreateObject("scripting.dictionary")
    Debug.Print TypeName(tes)
    
    'ArrayList
    Set tes = CreateObject("system.collections.arraylist")
    Debug.Print TypeName(tes)
    
    'Variant() -> Dimension, LBound, UBound
    dat = Range("A1:F12")
    Debug.Print TypeName(dat)
    
    Dim P()
    q = 5
    R = 3
    s = 6
    ReDim P(1 To q, 1 To R, 1 To s)
    dat = P
    Debug.Print TypeName(dat)
    
    'String() -> Dimension, LBound, UBound
    dat = Split("a,b,c,d,e", ",")
    Debug.Print TypeName(dat)
    
    'String
    dat = ","
    Debug.Print TypeName(dat)
    
    'Date
    dat = #5/1/2023#
    Debug.Print TypeName(dat)
    
    dat = #12:25:00 PM#
    Debug.Print TypeName(dat)
    
    
    'Long
    dat = 0&
    Debug.Print TypeName(dat)
    
    'Integer
    dat = 0
    Debug.Print TypeName(dat)
    
    
    'Double
    dat = 0#
    Debug.Print TypeName(dat)
    
    'Single
    dat = 99.9!
    Debug.Print TypeName(dat)

    
    'Empty
    Debug.Print TypeName(noexist)
    
    'Empty も代入可能
    dat = Empty
    Debug.Print TypeName(dat)
    
    'Nothing（Set文が必要）
    Set dat = Nothing
    Debug.Print TypeName(dat)
    
    'Null
    dat = Null
    Debug.Print TypeName(dat)

    'Object() や VARIANT() では要素ごとに型が違う可能性がある
    Dim A(1 To 6) As Object
    Set A(1) = CreateObject("scripting.dictionary")
    Set A(2) = CreateObject("system.collections.arraylist")
    Set A(3) = CreateObject("scripting.dictionary")
    Set A(4) = CreateObject("scripting.dictionary")
    Set A(5) = Nothing
    
    Debug.Print "Arr: " & TypeName(A)
    Debug.Print "1: " & TypeName(A(1))
    Debug.Print "2: " & TypeName(A(2))
    Debug.Print "3: " & TypeName(A(3))
    Debug.Print "4: " & TypeName(A(4))
    Debug.Print "5: " & TypeName(A(5))

    '★scripting.dictionaryのキーをCstrで強制的に文字列にするらっぱーとか
    
    dat = CStr("testtest")
    Debug.Print dat
    dat = CStr(12344)
    Debug.Print dat
    dat = CStr(#4/5/2023#)
    Debug.Print dat
    dat = CStr(#4:34:00 PM#)
    Debug.Print dat
    dat = CStr(#4/5/2023 12:23:00 PM#)
    Debug.Print dat
    
    
End Sub




Public Function GetDimension(ByRef ArrayData)
 
    temp = 0
    
    On Error Resume Next
    While Err.Number = 0
        temp = temp + 1
        dummy = UBound(ArrayData, temp)
    Wend
    On Error GoTo 0
    
    Err.Clear
    GetDimension = temp - 1
    
End Function

Private Function MultiDimArray2String(ByRef d) As String

    result = ""
    
    dimensions = GetDimension(d)

    For i = 1 To dimensions
        If result <> "" Then result = result + ","
        result = result + CStr(LBound(d, i)) + "," + CStr(UBound(d, i))
    Next

    MultiDimArray2String = CStr(dimensions) + "(" + result + ")"

End Function

Private Function String2MultiDimArray(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray = CreateMultiDimArray(vals)

End Function

'-- 自動ここから ------------------------------------------------------------------------
'-- 自動ここから ------------------------------------------------------------------------
'-- 自動ここから ------------------------------------------------------------------------
'-- 自動ここから ------------------------------------------------------------------------
'-- 自動ここから ------------------------------------------------------------------------

Private Function String2MultiDimArray_Variant(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Variant = CreateMultiDimArray_Variant(vals)

End Function

Private Function String2MultiDimArray_Object(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Object = CreateMultiDimArray_Object(vals)

End Function

Private Function String2MultiDimArray_String(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_String = CreateMultiDimArray_String(vals)

End Function

Private Function String2MultiDimArray_Byte(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Byte = CreateMultiDimArray_Byte(vals)

End Function

Private Function String2MultiDimArray_Long(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Long = CreateMultiDimArray_Long(vals)

End Function

Private Function String2MultiDimArray_Integer(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Integer = CreateMultiDimArray_Integer(vals)

End Function

Private Function String2MultiDimArray_Double(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Double = CreateMultiDimArray_Double(vals)

End Function

Private Function String2MultiDimArray_Single(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Single = CreateMultiDimArray_Single(vals)

End Function

Private Function String2MultiDimArray_Boolean(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Boolean = CreateMultiDimArray_Boolean(vals)

End Function

Private Function String2MultiDimArray_Date(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Date = CreateMultiDimArray_Date(vals)

End Function

Private Function String2MultiDimArray_Currency(ByRef sstr) As Variant

    str2 = Split(sstr, "(")
    dimensions = Val(str2(0))
    ReDim vals(0 To dimensions * 2 - 1)
    If Right(str2(1), 1) <> ")" Then Err.Raise Number:=700, Description:="エラーが発生！"
    valstr = Split(Left(str2(1), Len(str2(1)) - 1), ",")
    For i = 0 To dimensions * 2 - 1
        vals(i) = Val(valstr(i))
    Next

    String2MultiDimArray_Currency = CreateMultiDimArray_Currency(vals)

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がVariant()の場合
Private Function CreateMultiDimArray_Variant(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Variant
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Variant
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Variant
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Variant
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Variant
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Variant
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Variant
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Variant
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Variant
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Variant
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Variant
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Variant
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Variant
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Variant
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Variant
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Variant
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Variant
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Variant
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Variant
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Variant
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Variant
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Variant
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Variant
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Variant
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Variant = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がObject()の場合
Private Function CreateMultiDimArray_Object(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Object
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Object
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Object
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Object
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Object
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Object
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Object
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Object
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Object
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Object
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Object
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Object
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Object
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Object
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Object
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Object
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Object
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Object
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Object
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Object
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Object
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Object
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Object
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Object
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Object = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がString()の場合
Private Function CreateMultiDimArray_String(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As String
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As String
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As String
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As String
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As String
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As String
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As String
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As String
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As String
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As String
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As String
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As String
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As String
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As String
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As String
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As String
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As String
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As String
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As String
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As String
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As String
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As String
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As String
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As String
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_String = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がByte()の場合
Private Function CreateMultiDimArray_Byte(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Byte
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Byte
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Byte
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Byte
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Byte
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Byte
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Byte
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Byte
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Byte
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Byte
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Byte
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Byte
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Byte
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Byte
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Byte
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Byte
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Byte
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Byte
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Byte
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Byte
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Byte
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Byte
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Byte
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Byte
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Byte = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がLong()の場合
Private Function CreateMultiDimArray_Long(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Long
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Long
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Long
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Long
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Long
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Long
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Long
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Long
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Long
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Long
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Long
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Long
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Long
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Long
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Long
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Long
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Long
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Long
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Long
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Long
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Long
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Long
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Long
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Long
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Long = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がInteger()の場合
Private Function CreateMultiDimArray_Integer(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Integer
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Integer
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Integer
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Integer
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Integer
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Integer
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Integer
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Integer
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Integer
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Integer
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Integer
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Integer
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Integer
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Integer
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Integer
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Integer
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Integer
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Integer
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Integer
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Integer
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Integer
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Integer
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Integer
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Integer
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Integer = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がDouble()の場合
Private Function CreateMultiDimArray_Double(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Double
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Double
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Double
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Double
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Double
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Double
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Double
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Double
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Double
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Double
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Double
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Double
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Double
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Double
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Double
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Double
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Double
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Double
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Double
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Double
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Double
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Double
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Double
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Double
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Double = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がSingle()の場合
Private Function CreateMultiDimArray_Single(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Single
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Single
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Single
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Single
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Single
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Single
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Single
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Single
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Single
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Single
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Single
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Single
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Single
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Single
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Single
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Single
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Single
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Single
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Single
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Single
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Single
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Single
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Single
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Single
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Single = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がBoolean()の場合
Private Function CreateMultiDimArray_Boolean(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Boolean
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Boolean
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Boolean
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Boolean
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Boolean
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Boolean
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Boolean
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Boolean
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Boolean
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Boolean
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Boolean
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Boolean
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Boolean
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Boolean
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Boolean
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Boolean
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Boolean
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Boolean
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Boolean
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Boolean
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Boolean
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Boolean
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Boolean
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Boolean
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Boolean = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がDate()の場合
Private Function CreateMultiDimArray_Date(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Date
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Date
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Date
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Date
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Date
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Date
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Date
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Date
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Date
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Date
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Date
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Date
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Date
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Date
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Date
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Date
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Date
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Date
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Date
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Date
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Date
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Date
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Date
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Date
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Date = temp

End Function

'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること
'作る型がCurrency()の場合
Private Function CreateMultiDimArray_Currency(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK

    Dim temp

    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:="エラーが発生！"
    dimension = (UBound(bounds) + 1) \ 2
    Select Case dimension
    Case 0
        Err.Raise Number:=700, Description:="エラーが発生！"
    Case 1
        ReDim temp(bounds(0) To bounds(1)) As Currency
    Case 2
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3)) As Currency
    Case 3
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5)) As Currency
    Case 4
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7)) As Currency
    Case 5
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9)) As Currency
    Case 6
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11)) As Currency
    Case 7
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13)) As Currency
    Case 8
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15)) As Currency
    Case 9
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17)) As Currency
    Case 10
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19)) As Currency
    Case 11
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21)) As Currency
    Case 12
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23)) As Currency
    Case 13
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25)) As Currency
    Case 14
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27)) As Currency
    Case 15
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29)) As Currency
    Case 16
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31)) As Currency
    Case 17
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33)) As Currency
    Case 18
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35)) As Currency
    Case 19
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37)) As Currency
    Case 20
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39)) As Currency
    Case 21
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41)) As Currency
    Case 22
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43)) As Currency
    Case 23
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45)) As Currency
    Case 24
        ReDim temp(bounds(0) To bounds(1), bounds(2) To bounds(3), bounds(4) To bounds(5), bounds(6) To bounds(7), bounds(8) To bounds(9), bounds(10) To bounds(11), bounds(12) To bounds(13), bounds(14) To bounds(15), bounds(16) To bounds(17), bounds(18) To bounds(19), bounds(20) To bounds(21), bounds(22) To bounds(23), bounds(24) To bounds(25), bounds(26) To bounds(27), bounds(28) To bounds(29), bounds(30) To bounds(31), bounds(32) To bounds(33), bounds(34) To bounds(35), bounds(36) To bounds(37), bounds(38) To bounds(39), bounds(40) To bounds(41), bounds(42) To bounds(43), bounds(44) To bounds(45), bounds(46) To bounds(47)) As Currency
    Case Else
        Err.Raise Number:=700, Description:="エラーが発生！"
    End Select

    CreateMultiDimArray_Currency = temp

End Function

'-- 自動ここまで ------------------------------------------------------------------------
'-- 自動ここまで ------------------------------------------------------------------------
'-- 自動ここまで ------------------------------------------------------------------------
'-- 自動ここまで ------------------------------------------------------------------------
'-- 自動ここまで ------------------------------------------------------------------------


Private Sub test_createmultidimarray()

    B = Array(0, 4, 2, 6)
    result = CreateMultiDimArray(B)

    Debug.Print LBound(result, 1)

    ss = MultiDimArray2String(result)
    Debug.Print ss

    result2 = String2MultiDimArray(ss)
    ss = MultiDimArray2String(result2)
    Debug.Print ss

End Sub



    Private Sub SerializeRecurWithLoop(ByRef result, ByRef v)
        dimensions = GetDimension(v)
        ReDim lbounds(1 To dimensions)
        ReDim ubounds(1 To dimensions)
    
        For i = 1 To dimensions
            lbounds(i) = LBound(v, i)
            ubounds(i) = UBound(v, i)
        Next
        
        Select Case dimensions
        Case 1
            For i1 = LBound(v, 1) To UBound(v, 1)
                SerializeRecur result, v(i1)
            Next
        Case 2
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
                SerializeRecur result, v(i1, i2)
            Next
            Next
        Case 3
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
                SerializeRecur result, v(i1, i2, i3)
            Next
            Next
            Next
        Case 4
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
                SerializeRecur result, v(i1, i2, i3, i4)
            Next
            Next
            Next
            Next
        Case 5
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
                SerializeRecur result, v(i1, i2, i3, i4, i5)
            Next
            Next
            Next
            Next
            Next
        Case 6
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6)
            Next
            Next
            Next
            Next
            Next
            Next
        Case 7
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 8
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 9
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 10
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 11
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 12
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 13
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 14
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 15
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 16
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 17
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 18
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 19
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
            For i19 = LBound(v, 19) To UBound(v, 19)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 20
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
            For i19 = LBound(v, 19) To UBound(v, 19)
            For i20 = LBound(v, 20) To UBound(v, 20)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 21
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
            For i19 = LBound(v, 19) To UBound(v, 19)
            For i20 = LBound(v, 20) To UBound(v, 20)
            For i21 = LBound(v, 21) To UBound(v, 21)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 22
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
            For i19 = LBound(v, 19) To UBound(v, 19)
            For i20 = LBound(v, 20) To UBound(v, 20)
            For i21 = LBound(v, 21) To UBound(v, 21)
            For i22 = LBound(v, 22) To UBound(v, 22)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21, i22)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 23
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
            For i19 = LBound(v, 19) To UBound(v, 19)
            For i20 = LBound(v, 20) To UBound(v, 20)
            For i21 = LBound(v, 21) To UBound(v, 21)
            For i22 = LBound(v, 22) To UBound(v, 22)
            For i23 = LBound(v, 23) To UBound(v, 23)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21, i22, i23)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 24
            For i1 = LBound(v, 1) To UBound(v, 1)
            For i2 = LBound(v, 2) To UBound(v, 2)
            For i3 = LBound(v, 3) To UBound(v, 3)
            For i4 = LBound(v, 4) To UBound(v, 4)
            For i5 = LBound(v, 5) To UBound(v, 5)
            For i6 = LBound(v, 6) To UBound(v, 6)
            For i7 = LBound(v, 7) To UBound(v, 7)
            For i8 = LBound(v, 8) To UBound(v, 8)
            For i9 = LBound(v, 9) To UBound(v, 9)
            For i10 = LBound(v, 10) To UBound(v, 10)
            For i11 = LBound(v, 11) To UBound(v, 11)
            For i12 = LBound(v, 12) To UBound(v, 12)
            For i13 = LBound(v, 13) To UBound(v, 13)
            For i14 = LBound(v, 14) To UBound(v, 14)
            For i15 = LBound(v, 15) To UBound(v, 15)
            For i16 = LBound(v, 16) To UBound(v, 16)
            For i17 = LBound(v, 17) To UBound(v, 17)
            For i18 = LBound(v, 18) To UBound(v, 18)
            For i19 = LBound(v, 19) To UBound(v, 19)
            For i20 = LBound(v, 20) To UBound(v, 20)
            For i21 = LBound(v, 21) To UBound(v, 21)
            For i22 = LBound(v, 22) To UBound(v, 22)
            For i23 = LBound(v, 23) To UBound(v, 23)
            For i24 = LBound(v, 24) To UBound(v, 24)
                SerializeRecur result, v(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21, i22, i23, i24)
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case Else
            Err.Raise Number:=700, Description:="エラーが発生！"
        End Select
        
    End Sub
    Private Sub DeserializeRecurWithLoop(ByRef obj, ByRef token, ByRef ind)
        'もう配列の入れ物だけは出来上がっている前提
        dimensions = GetDimension(obj)
        ReDim lbounds(1 To dimensions)
        ReDim ubounds(1 To dimensions)
    
        For i = 1 To dimensions
            lbounds(i) = LBound(obj, i)
            ubounds(i) = UBound(obj, i)
        Next
        
        Select Case dimensions
        Case 1
            For i1 = LBound(obj, 1) To UBound(obj, 1)
                DeserializeRecur obj(i1), token, ind
            Next
        Case 2
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
                DeserializeRecur obj(i1, i2), token, ind
            Next
            Next
        Case 3
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
                DeserializeRecur obj(i1, i2, i3), token, ind
            Next
            Next
            Next
        Case 4
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
                DeserializeRecur obj(i1, i2, i3, i4), token, ind
            Next
            Next
            Next
            Next
        Case 5
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
                DeserializeRecur obj(i1, i2, i3, i4, i5), token, ind
            Next
            Next
            Next
            Next
            Next
        Case 6
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
        Case 7
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 8
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 9
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 10
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 11
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 12
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 13
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 14
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 15
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 16
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 17
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 18
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 19
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
            For i19 = LBound(obj, 19) To UBound(obj, 19)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 20
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
            For i19 = LBound(obj, 19) To UBound(obj, 19)
            For i20 = LBound(obj, 20) To UBound(obj, 20)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 21
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
            For i19 = LBound(obj, 19) To UBound(obj, 19)
            For i20 = LBound(obj, 20) To UBound(obj, 20)
            For i21 = LBound(obj, 21) To UBound(obj, 21)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 22
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
            For i19 = LBound(obj, 19) To UBound(obj, 19)
            For i20 = LBound(obj, 20) To UBound(obj, 20)
            For i21 = LBound(obj, 21) To UBound(obj, 21)
            For i22 = LBound(obj, 22) To UBound(obj, 22)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21, i22), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 23
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
            For i19 = LBound(obj, 19) To UBound(obj, 19)
            For i20 = LBound(obj, 20) To UBound(obj, 20)
            For i21 = LBound(obj, 21) To UBound(obj, 21)
            For i22 = LBound(obj, 22) To UBound(obj, 22)
            For i23 = LBound(obj, 23) To UBound(obj, 23)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21, i22, i23), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case 24
            For i1 = LBound(obj, 1) To UBound(obj, 1)
            For i2 = LBound(obj, 2) To UBound(obj, 2)
            For i3 = LBound(obj, 3) To UBound(obj, 3)
            For i4 = LBound(obj, 4) To UBound(obj, 4)
            For i5 = LBound(obj, 5) To UBound(obj, 5)
            For i6 = LBound(obj, 6) To UBound(obj, 6)
            For i7 = LBound(obj, 7) To UBound(obj, 7)
            For i8 = LBound(obj, 8) To UBound(obj, 8)
            For i9 = LBound(obj, 9) To UBound(obj, 9)
            For i10 = LBound(obj, 10) To UBound(obj, 10)
            For i11 = LBound(obj, 11) To UBound(obj, 11)
            For i12 = LBound(obj, 12) To UBound(obj, 12)
            For i13 = LBound(obj, 13) To UBound(obj, 13)
            For i14 = LBound(obj, 14) To UBound(obj, 14)
            For i15 = LBound(obj, 15) To UBound(obj, 15)
            For i16 = LBound(obj, 16) To UBound(obj, 16)
            For i17 = LBound(obj, 17) To UBound(obj, 17)
            For i18 = LBound(obj, 18) To UBound(obj, 18)
            For i19 = LBound(obj, 19) To UBound(obj, 19)
            For i20 = LBound(obj, 20) To UBound(obj, 20)
            For i21 = LBound(obj, 21) To UBound(obj, 21)
            For i22 = LBound(obj, 22) To UBound(obj, 22)
            For i23 = LBound(obj, 23) To UBound(obj, 23)
            For i24 = LBound(obj, 24) To UBound(obj, 24)
                DeserializeRecur obj(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19, i20, i21, i22, i23, i24), token, ind
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
            Next
        Case Else
            Err.Raise Number:=700, Description:="エラーが発生！"
        End Select
        
    End Sub
    
    
    
'上のコードは自動生成しないと作れないので
Private Sub createcode_CreateMultiDimArray()

    For N = 1 To 24
        Debug.Print "    Case " & N
        txt = "bounds(0) To bounds(1)"
        For i = 2 To N * 2 - 2 Step 2
            txt = txt + ", bounds(" & CStr(i) & ") To bounds(" & CStr(i + 1) & ")"
        Next
        Debug.Print "        ReDim temp(" & txt & ")"
    Next

End Sub

Private Sub createcode_CreateMultiDimArray_AS_XXX()

    strarr = Split("Variant,Object,String,Byte,Long,Integer,Double,Single,Boolean,Date,Currency", ",")

    Debug.Print "   '配列のシリアライズはすべて同じ形"
    Debug.Print "        '型名、次元数（要素数列挙）、値×Π要素数列挙（SerializeRecurWithLoop）"

    For J = LBound(strarr) To UBound(strarr)
        tn = strarr(J)
        Debug.Print "        Case """ & tn & "()"""
        Debug.Print "            result = result + TypeName(v) + ""|"""
        Debug.Print "            result = result + MultiDimArray2String(v) + ""|"""
        Debug.Print "            SerializeRecurWithLoop result, v"
    Next
    
    Debug.Print ""
        
    Debug.Print "   '配列のデシリアライズはすべて同じ形"
    Debug.Print "        '型名、次元数（要素数列挙）、値×Π要素数列挙（DeserializeRecurWithLoop）"

    For J = LBound(strarr) To UBound(strarr)
        tn = strarr(J)
        Debug.Print "        Case """ & tn & "()"""
        Debug.Print "            obj = String2MultiDimArray_" & tn & "(token(ind))"
        Debug.Print "            ind = ind + 1"
        Debug.Print "            DeserializeRecurWithLoop obj, token, ind"
    Next

    Open "D:\Users\miyokomizo\Desktop\test\srcc.txt" For Output As #1
    
    For J = LBound(strarr) To UBound(strarr)
        tn = strarr(J)
        
        Print #1, "Private Function String2MultiDimArray_" & tn & "(ByRef sstr) As Variant"
        Print #1, ""
        Print #1, "    str2 = Split(sstr, ""("")"
        Print #1, "    dimensions = Val(str2(0))"
        Print #1, "    ReDim vals(0 To dimensions * 2 - 1)"
        Print #1, "    If Right(str2(1), 1) <> "")"" Then Err.Raise Number:=700, Description:=""エラーが発生！"""
        Print #1, "    valstr = Split(Left(str2(1), Len(str2(1)) - 1), "","")"
        Print #1, "    For i = 0 To dimensions * 2 - 1"
        Print #1, "        vals(i) = Val(valstr(i))"
        Print #1, "    Next"
        Print #1, ""
        Print #1, "    String2MultiDimArray_" & tn & " = CreateMultiDimArray_" & tn & "(vals)"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    Next

    For J = LBound(strarr) To UBound(strarr)
        tn = strarr(J)

        Print #1, "'n次元配列を作る。bounds(0...dimensions*2-1)に、LBound, UBoundのリストを入れること"
        Print #1, "'作る型が" & tn & "()の場合"
        Print #1, "Private Function CreateMultiDimArray_" & tn & "(ByRef bounds) As Variant 'ここは配列をVariantに収めて返すのでVariantでOK"
        Print #1, ""
        Print #1, "    Dim temp"
        Print #1, ""
        Print #1, "    If LBound(bounds) <> 0 Then Err.Raise Number:=700, Description:=""エラーが発生！"""
        Print #1, "    dimension = (UBound(bounds) + 1) \ 2"
        Print #1, "    Select Case dimension"
        Print #1, "    Case 0"
        Print #1, "        Err.Raise Number:=700, Description:=""エラーが発生！"""
    
        For N = 1 To 24
            Print #1, "    Case " & N
            txt = "bounds(0) To bounds(1)"
            For i = 2 To N * 2 - 2 Step 2
                txt = txt + ", bounds(" & CStr(i) & ") To bounds(" & CStr(i + 1) & ")"
            Next
            Print #1, "        ReDim temp(" & txt & ") As " & tn
        Next
    
        Print #1, "    Case Else"
        Print #1, "        Err.Raise Number:=700, Description:=""エラーが発生！"""
        Print #1, "    End Select"
        Print #1, ""
        Print #1, "    CreateMultiDimArray_" & tn & " = temp"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    Next


    Close #1

End Sub

Private Sub createcode_SerializeRecurWithLoop()

    Open "D:\Users\miyokomizo\Desktop\test\srcs.txt" For Output As #1
    For dimensions = 1 To 24
        Print #1, "        Case " & dimensions
    
        txt = ""
        For i = 1 To dimensions
            Print #1, "            For i" & i & " = LBound(v, " & i & ") To UBound(v, " & i & ")"
            If txt <> "" Then txt = txt + ", "
            txt = txt + "i" + CStr(i)
        Next
        Print #1, "                SerializeRecur result, v(" & txt & ")"
        For i = 1 To dimensions
            Print #1, "            Next"
        Next
    Next
    Close #1

End Sub

Private Sub createcode_DeserializeRecurWithLoop()

    Open "D:\Users\miyokomizo\Desktop\test\srcd.txt" For Output As #1
    For dimensions = 1 To 24
        Print #1, "        Case " & dimensions
    
        txt = ""
        For i = 1 To dimensions
            Print #1, "            For i" & i & " = LBound(obj, " & i & ") To UBound(obj, " & i & ")"
            If txt <> "" Then txt = txt + ", "
            txt = txt + "i" + CStr(i)
        Next
        Print #1, "                DeserializeRecur obj(" & txt & "), token, ind"
        For i = 1 To dimensions
            Print #1, "            Next"
        Next
    Next

End Sub

