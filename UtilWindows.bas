Attribute VB_Name = "UtilWindows"
'=========================================================================================
'UtilWindows 20230527
'
'UtilWindowsは主にWindows OSの制御を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'クリップボードにプレーンテキストをセットする
'Public Sub SetClip(txt)
'クリップボードからプレーンテキストを読み取る
'Public Function GetClip()
'URLとメソッドを指定してウェブにアクセスし、結果を文字列で得る
'Public Function HostApplication()
'このマクロがどのアプリに組み込まれているか"Microsoft Excel"とかで分岐するため
'Public Function EscapedSplit(txt, delim)
'多次元配列の次元数を得る
'Public Function GetDimension(ByRef ArrayData)
'=========================================================================================
'https://gist.github.com/KotorinChunChun/718da75c26de71c9e4b12afa9c19ee32
Type coord
    x As Long
    y As Long
End Type
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
        'SetCursorPos　・・・マウスを動かす・マウスのポインターの操作を行う。
        Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
        'Mouseevent　　・・・マウスをクリックする操作を行う。
        Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, Optional ByVal dx As Long = 0, Optional ByVal dy As Long = 0, Optional ByVal dwDate As Long = 0, Optional ByVal dwExtraInfo As Long = 0)
        'GetCursorPos　・・・マウスのポインターの位置を取得します。
        Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As coord) As Long
    #Else
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
        'SetCursorPos　・・・マウスを動かす・マウスのポインターの操作を行う。
        Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
        'Mouseevent　　・・・マウスをクリックする操作を行う。
        Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, Optional ByVal dx As Long = 0, Optional ByVal dy As Long = 0, Optional ByVal dwDate As Long = 0, Optional ByVal dwExtraInfo As Long = 0)
        'GetCursorPos　・・・マウスのポインターの位置を取得します。
        Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As coord) As Long
    #End If
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If



'重要！Excel以外でも共通コードで実行するためには、下記のようにすべし
Private Sub test_multi_host()

    '重要！Excel以外で実行できないコードの実行を防ぐ
    If Application.Name = "Microsoft Excel" Then
        Debug.Print ThisWorkbook.Name
        'サブルーチン／関数単位のコンパイル時のエラーを防ぐため、Applicationを参照経由で叩く
        Set xlapp = Application
        'サブルーチン／関数単位のコンパイル時のエラーを防ぐため、Applicationを参照経由で叩く
        xlapp.ActiveSheet.Range("A2") = "aaa"
        xlapp.ActiveSheet.Range("A2").Clear
        '省略記法がつかえないだけで、割合自然に書ける
    End If

End Sub




'クリップボード処理

Public Sub SetClip(txt)
    'http://www.thom.jp/vbainfo/refsetting.html
    Set dao = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dao.settext txt
    dao.PutInClipboard
End Sub

Public Function GetClip()
    'http://www.thom.jp/vbainfo/refsetting.html
    Set dao = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dao.GetFromClipboard
    
    Set flag = CreateObject("scripting.dictionary")
    fmt = Application.ClipboardFormats
    For i = LBound(fmt) To UBound(fmt)
        flag.Add fmt(i), i
        Debug.Print fmt(i)
    Next
'0: テキスト
'2: 画像
'9: BitMap
'47: ファイルパス
'14:画像系？
'17:画像系？
'22:画像系？
'31:画像系？
'45:画像系？
    
    If flag.exists(0) Then
        GetClip = dao.GetText
    Else
        GetClip = ""
    End If
    
    'GetClip = dao.GetImage
End Function



'ウェブアクセス（これはAPI叩く用）
'他にEdgeのDOMを覗く方法もあるらしい

Public Function Web(url, method)

    Set xmlhttp = CreateObject("msxml2.xmlhttp")
    xmlhttp.Open method, url
    xmlhttp.Send
    
    Do While xmlhttp.ReadyState < 4
        DoEvents
    Loop
    
    Web = xmlhttp.responseText

End Function

Sub test_web()

    Debug.Print Web("https://www.google.com/", "GET")

End Sub



'オートパイロット系

Private Sub test_autoit()

    'SendKeys "test"

End Sub


Private Sub test_multihost_function()

    If HostApplication = "Microsoft PowerPoint" Then
        Set ppApp = Application
    Else
        Set ppApp = CreateObject("PowerPoint.Application")
    End If
    
    If HostApplication = "Microsoft Excel" Then
        Set xlapp = Application
    Else
        Set xlapp = CreateObject("Excel.Application")
    End If
    
    If HostApplication = "Microsoft Word" Then
        Set wdapp = Application
    Else
        Set wdapp = CreateObject("Word.Application")
    End If
    
End Sub


Public Function HostApplication()

    Debug.Print Application.Name
    HostApplication = Application.Name

End Function



