Attribute VB_Name = "UtilMath"
'=========================================================================================
'UtilMath 20230527
'
'UtilMathは主にファイル操作を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'UTF8のファイルをSJISに変換する
'Sub Utf8ToSjis(a_sFrom, a_sTo)
'SJISのファイルをUTF8に変換する
'Sub SjisToUtf8(a_sFrom, a_sTo)
'GUIDを生成する
'Public Function GetGUID()
'フォルダを下位のファイルやサブディレクトリ含め、可能な限り削除する
'Function RmDirBestEffort(ByVal sDir As String, ByRef sMsg As String, Optional ByVal isOnlyFile As Boolean = False) As Boolean
'ファイルリストを得る
'Public Function GetFileList(ByVal path, Optional ext = "")
'ファイルリストを得る（ファイル名を正規表現で指定する）
'Public Function GetFileListRegex(ByVal path, Optional ByVal recur = False, Optional pat = ".*")
'フォルダ選択ダイアログ
'Public Function FolderPicker(defpath)
'ファイル名に良く仕込む時刻文字列
'Public Function TimeString()
'=========================================================================================

Function ArcSin(x As Double) As Double

　　Select Case x
　　　　Case Is = -1#
　　　　ArcSin = -1.57079632679
　　　　Case Is = 1#
　　　　ArcSin = 1.57079632679
　　　　Case Else
　　　　ArcSin = Atn(x / Sqr(1 - x * x))
　　End Select

End Function

Function ArcCos(x As Double) As Double

　　Select Case x
　　　　Case Is = -1#
　　　　ArcCos = 3.1415926535897932
　　　　Case Is = 1#
　　　　ArcCos = 0
　　　　Case Else
　　　　ArcCos = 1.57079632679 - Atn(x / Sqr(1 - x * 2))
　　End Select

End Function

Function Curvature(x As Double, y As Double, x_prev As Double, y_prev As Double, x_next As Double, y_next As Double) As Double

    x1 = x_prev - x
    y1 = y_prev - y
    div1 = 1# / Sqr(x1 * x1 + y1 * y1)
    x1 = x1 * div1
    y1 = y1 * div1
    x2 = x_next - x
    y2 = y_next - y
    div2 = 1# / Sqr(x2 * x2 + y2 * y2)
    x2 = x2 * div2
    y2 = y2 * div2
    innerp = x1 * x2 + y1 * y2
    outerp = x1 * y2 - y1 * x2
    signp = Sgn(outerp)
    If innerp > 0.99999999999999 Then
        Curvature = signp * 1E+20
        Exit Function
    End If
    Curvature = signp * 2# * Sqr((1 + innerp) / (1 - innerp))

End Function

Sub test_Curvature()

    Debug.Print Curvature(0, 0, 1, 0, 0, 1) & " -> clockwise"
    Debug.Print Curvature(1, 1, 1, 0, 0, 1) & " -> anti->clockwise"
    

End Sub
