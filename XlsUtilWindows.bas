Attribute VB_Name = "XlsUtilWindows"
'=========================================================================================
'XlsUtilWindows 20230527
'
'XlsUtilWindowsは主にWindows OSの制御を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'クリップボードにプレーンテキストをセットする
'Public Sub SetClip(txt)
'クリップボードからプレーンテキストを読み取る
'Public Function GetClip()
'=========================================================================================
'https://gist.github.com/KotorinChunChun/718da75c26de71c9e4b12afa9c19ee32


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

