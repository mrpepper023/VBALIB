Attribute VB_Name = "XlsUtilSheet"
Const DANGERFAST = True
'=========================================================================================
'XlsUtilSheet 20230527
'
'XlsUtilSheetは
'=========================================================================================
'ワークシート向けに、半角変換する関数
'Public Function HanKana(str)
'シートオブジェクトと列番号を渡し、該当列の最終行の行番号を取得する
'Public Function GetMaxRow(ByRef sh, Optional ByVal col = 1)
'シートオブジェクトと行番号を渡し、該当行の最終列の列番号を取得する
'Public Function GetMaxCol(ByRef sh, Optional ByVal row = 1)
'シートオブジェクトを渡し、使用した最終行の行番号を取得する
'Public Function GetUsedMaxRow(ByRef sh)
'シートオブジェクトを渡し、使用した最終列の列番号を取得する
'Public Function GetUsedMaxCol(ByRef sh)
'シートオブジェクトと左上右下の直値を元にしてRangeオブジェクトを返す
'Public Function RectRange(ByRef sheet, ByVal r_top, ByVal c_left, Optional ByVal r_bottom = 0, Optional ByVal c_right = 0)
'Rangeオブジェクトを渡して、手編集セルの装飾を行う
'Public Sub DecorateManualCells(ByRef rng, Optional ByVal defval = "")
'Excelマクロの単純な高速化：エラー処理が面倒になるので、バグがとれてから仕込もう。
'Public Sub FastSetting(flag)
'ビジーループがあるなら、ループ内でこれを呼んでおけ
'Public Sub FastSettingDoEvents(Optional str = "")
'どうしても止めたい場合、これを使う
'Public Sub FastSettingStop(Optional str = "")
'コメントをStatusbarとImmediate Windowに表示
'Public Sub d_print_____(str)
'A,B,AA,ABなどの列表記のアルファベットを、1から順に振った列番号に変換する
'Public Function addr2col(ByVal str)
'1から順に振った列番号を、A,B,AA,ABなどの列表記のアルファベットに変換する
'Public Function col2addr(ByVal num)
'targetstrの文字列の冒頭が、prefixになっているかどうかを判定する
'Public Function StartsWith(ByRef targetstr, ByRef prefix)
'targetstrの文字列の末尾が、suffixになっているかどうかを判定する
'Public Function EndsWith(ByRef targetstr, ByRef suffix)
'targetstrの先頭から、prefixを削除する。先頭が一致していなければ、targetstrをそのまま返す
'Public Function RemovePrefix(ByRef targetstr, ByRef prefix)
'targetstrの末尾から、suffixを削除する。末尾が一致していなければ、targetstrをそのまま返す
'Public Function RemoveSuffix(ByRef targetstr, ByRef suffix)
'Excel専用！指定したファイルが開かれていればそのブックを、さもなくばファイルを開いてブックを得る
'Public Function WiseOpen(path, ByRef closeflag)
'引数で示した名前のシートを探して、存在するならTRUE、存在しないならFALSEを返す
'Public Function FindSheet(ByVal str)
'引数で示した名前のシートを探して、存在しないならエラーメッセージを表示して中断する
'Public Sub FindSheet_Trap(ByVal str)
'★★FindSheetRegexも作っておきたい
'prefで指定した文字列で始まる名前のシートを、配列変数に格納して返す
'Public Function FindSheetPrefix(ByVal pref)
'シートへの参照と文字列（と行番号）を指定して、特定の文字列と一致する表題のついた列を見つけて返す（見つからない場合は0を返す
'Public Function FindCol(ByRef sheet, ByRef str, Optional ByVal row = 1)
'★★シートの表題行の推測も作っておきたい
'★★列特性取得とか、あってもいい。
'=========================================================================================




'ワークシート向けに、半角変換する関数
Public Function HanKana(str)

    HanKana = StrConv(str, vbNarrow)
    
End Function

'シートオブジェクトと列番号を渡し、該当列の最終行の行番号を取得する
Public Function GetMaxRow(ByRef sh, Optional ByVal col = 1)
    
    GetMaxRow = sh.Cells(sh.Cells(sh.Rows.Count, col).row, col).End(xlUp).row
    
End Function

'シートオブジェクトと行番号を渡し、該当行の最終列の列番号を取得する
Public Function GetMaxCol(ByRef sh, Optional ByVal row = 1)
    
    GetMaxCol = sh.Cells(row, sh.Cells(row, sh.Columns.Count).col).End(xlToLeft).Column
    
End Function 'シートオブジェクトを渡し、使用した最終行の行番号を取得する

'シートオブジェクトを渡し、使用した最終行の行番号を取得する
Public Function GetUsedMaxRow(ByRef sh)
    
    GetUsedMaxRow = sh.UsedRange.Rows(sh.UsedRange.Rows.Count).row
    
End Function

'シートオブジェクトを渡し、使用した最終列の列番号を取得する
Public Function GetUsedMaxCol(ByRef sh)
    
    GetUsedMaxCol = sh.UsedRange.Columns(sh.UsedRange.Columns.Count).Column
    
End Function



'シートオブジェクトと左上右下の直値を元にしてRangeオブジェクトを返す
Public Function RectRange(ByRef sheet, ByVal r_top, ByVal c_left, Optional ByVal r_bottom = 0, Optional ByVal c_right = 0)

    If r_bottom = 0 Then r_bottom = r_top
    If c_right = 0 Then c_right = c_left
    With sheet
        Set RectRange = .Range(.Cells(r_top, c_left), .Cells(r_bottom, c_right))
    End With
    
End Function



'Rangeオブジェクトを渡して、手編集セルの装飾を行う
Public Sub DecorateManualCells(ByRef rng, Optional ByVal defval = "")
    
    If Len(defval) > 0 Then
        rng.Value = defval
    End If
    
    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 0)
    End With
    
End Sub





'Excelマクロの単純な高速化：エラー処理が面倒になるので、バグがとれてから仕込もう。
Public Sub FastSetting(flag)
    If DANGERFAST Then
        If flag Then
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
        Else
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Application.Calculation = xlCalculationAutomatic
        End If
    End If
End Sub

'ビジーループがあるなら、ループ内でこれを呼んでおけ
Public Sub FastSettingDoEvents(Optional str = "")
    
    If Rnd < 0.01 Then
        savedsu = Application.ScreenUpdating
        savedee = Application.EnableEvents
        savedcc = Application.Calculation
        
        FastSetting False
        
        Application.StatusBar = str
        Debug.Print "[" & Format(Time(), "yyyy/mm/dd hh:nn:ss") & "] " & str
        DoEvents
        
        Application.ScreenUpdating = savedsu
        Application.EnableEvents = savedee
        Application.Calculation = savedcc
    End If
    
End Sub

'どうしても止めたい場合、これを使う
Public Sub FastSettingStop(Optional str = "")
    
    If Rnd < 0.01 Then
        savedsu = Application.ScreenUpdating
        savedee = Application.EnableEvents
        savedcc = Application.Calculation
        
        FastSetting False
        
        Application.StatusBar = str
        Debug.Print "[" & Format(Time(), "yyyy/mm/dd hh:nn:ss") & "] " & str
        Stop
        
        Application.ScreenUpdating = savedsu
        Application.EnableEvents = savedee
        Application.Calculation = savedcc
    End If
    
End Sub





' コメントをStatusbarとImmediate Windowに表示
Public Sub d_print_____(str)
    savedsu = Application.ScreenUpdating
    savedee = Application.EnableEvents
    savedcc = Application.Calculation
    FastSetting False

    Application.StatusBar = str
    Debug.Print "[" & Format(Time(), "yyyy/mm/dd hh:nn:ss") & "] " & str
    DoEvents
    
    Application.ScreenUpdating = savedsu
    Application.EnableEvents = savedee
    Application.Calculation = savedcc
End Sub



' A,B,AA,ABなどの列表記のアルファベットを、1から順に振った列番号に変換する
Public Function addr2col(ByVal str)
    On Error GoTo e
    addr2col = ActiveSheet.Range(str & 1).Column
    On Error GoTo 0
    Exit Function
e:
    On Error GoTo 0
    addr2col = False
End Function

'1から順に振った列番号を、A,B,AA,ABなどの列表記のアルファベットに変換する
Public Function col2addr(ByVal num)
    On Error GoTo e
    col2addr = Split(ActiveSheet.Cells(1, num).Address(True, False), "$")(0)
    On Error GoTo 0
    Exit Function
e:
    On Error GoTo 0
    col2addr = False
End Function



Private Sub test_addr2col()
    Debug.Print addr2col("A")
    Debug.Print addr2col("AA")
    Debug.Print addr2col("Xfd")
    Debug.Print addr2col("XXX")
    Debug.Print col2addr(2)
    Debug.Print col2addr(23)
    Debug.Print col2addr(235)
    Debug.Print col2addr(16251)
    Debug.Print col2addr(46251)
End Sub



'targetstrの文字列の冒頭が、prefixになっているかどうかを判定する
' 「京浜急行」「京浜」ならTRUE「焼肉定食」「肉食」ならFALSE
Public Function StartsWith(ByRef targetstr, ByRef prefix)
    StartsWith = (Left(targetstr, Len(prefix)) = prefix)
End Function
'targetstrの文字列の末尾が、suffixになっているかどうかを判定する
Public Function EndsWith(ByRef targetstr, ByRef suffix)
    EndsWith = (Right(targetstr, Len(suffix)) = suffix)
End Function

'targetstrの先頭から、prefixを削除する。先頭が一致していなければ、targetstrをそのまま返す
' 「京浜急行」「京浜」なら返値は「急行」
Public Function RemovePrefix(ByRef targetstr, ByRef prefix)
    If StartsWith(targetstr, prefix) Then
        RemovePrefix = Right(targetstr, Len(targetstr) - Len(prefix))
        Exit Function
    End If
    RemovePrefix = targetstr
End Function
' targetstrの末尾から、suffixを削除する。末尾が一致していなければ、targetstrをそのまま返す
Public Function RemoveSuffix(ByRef targetstr, ByRef suffix)
    If EndsWith(targetstr, suffix) Then
        RemoveSuffix = Left(targetstr, Len(targetstr) - Len(suffix))
        Exit Function
    End If
    RemoveSuffix = targetstr
End Function



'賢いファイルオープン（既に開いていればそちらを参照）してブックオブジェクトを返す
'使用後、closeflagを参照してTrueならファイルを閉じるのを推奨する
'なお、WiseOpenは基本的に読み取り専用で開く（書込対象をこういう仕様で選ぶのはあぶねー）
'なお、Excel専用
Public Function WiseOpen(path, ByRef closeflag)
    Set fso = CreateObject("scripting.filesystemobject")
    closeflag = False
    
    fname = fso.GetFileName(path)
    For Each f In Workbooks
        If fname = f.Name Then
            Set WiseOpen = f
            GoTo exitwiseopen
        End If
    Next
    
    closeflag = True
    savedupdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set WiseOpen = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
    Application.Windows(fname).Visible = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedupdating
    
exitwiseopen:
    Set fso = Nothing
End Function

Private Sub testwiseopen()

    Set bk = WiseOpen("D:\Users\miyokomizo\Desktop\編集：製造原価仕訳日記帳・借方（2021年3月1日～2022年2月28日）.xlsx", closeflag)
    If bk Is Nothing Then End
    
    If closeflag Then bk.Close savechanges:=False
    Set bk = Nothing

End Sub





' 引数で示した名前のシートを探して、存在するならTRUE、存在しないならFALSEを返す
Public Function FindSheet(ByVal str)
    For Each N In ThisWorkbook.Sheets
        If N.Name = str Then
            FindSheet = True
            Exit Function
        End If
    Next
    
    FindSheet = False
End Function

' 引数で示した名前のシートを探して、存在しないならエラーメッセージを表示して中断する
Public Sub FindSheet_Trap(ByVal str)

    If Not FindSheet(str) Then
        MsgBox str & "が見つかりません"
        End
    End If

End Sub

Private Sub test_findsheet()

    FindSheet_Trap "RAW1"
    FindSheet_Trap "RAW2"
    FindSheet_Trap "加工済データ集計設定"
    FindSheet_Trap "RAW1処理済"
    FindSheet_Trap "RAW2処理済"

End Sub



' prefで指定した文字列で始まる名前のシートを、配列変数に格納して返す
' 列挙する際にはlbound,uboundを使うこと
Public Function FindSheetPrefix(ByVal pref)
    Set temp = CreateObject("system.collections.arraylist")
    
    preflen = Len(pref)
    For Each N In ThisWorkbook.Sheets
        If Left(N.Name, preflen) = pref Then
            temp.Add N.Name
        End If
    Next
    
    FindSheetPrefix = temp.ToArray()
    Set temp = Nothing

End Function

Private Sub test_FindSheetPrefix()

    a = FindSheetPrefix("RAW")
    For i = LBound(a) To UBound(a)
        Debug.Print a(i)
    Next

End Sub



' シートへの参照と文字列（と行番号）を指定して、特定の文字列と一致する表題のついた列を見つけて返す
' 見つからない場合は0を返す
Public Function FindCol(ByRef sheet, ByRef str, Optional ByVal row = 1)
    maxc = MaxColumn(sheet, row)
    For i = 1 To maxc
        If str = sheet.Cells(row, i) Then
            FindCol = i
            Exit Function
        End If
    Next
    FindCol = 0
End Function

Private Sub test_FindCol()

    Debug.Print FindCol(ThisWorkbook.Sheets("加工済データ集計設定"), "テスト設問")
    Debug.Print FindCol(ThisWorkbook.Sheets("加工済データ集計設定"), "存在しないよ")

End Sub







