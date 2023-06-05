Attribute VB_Name = "PortMacro"
'=========================================================================================
'PortMacro 20230604
'
'UtilFileは主にファイル操作を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'GitHubからダウンロードしたzipからzip-extract.batで前処理したファイルをインポートする
'Public Sub MyImportTask()
'GitHubへエクスポートする（特定マシン専用）
'Public Sub MyExportTask()
'指定フォルダからマクロをインポートする
'Public Sub ImportAll(a_sModulePath)
'指定フォルダからマクロをエクスポートする
'Public Sub ExportAll(sPath)
'=========================================================================================
Private Type GUID_TYPEW
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr


Private Sub MyUtf8ToSjis(a_sFrom, a_sTo)
    Dim sText                           '// ファイルデータ
    Set streamRead = CreateObject("ADODB.Stream")
    Set streamWrite = CreateObject("ADODB.Stream")
    
    '// ファイル読み込み
    streamRead.Type = 2 'adTypeText
    streamRead.Charset = "UTF-8"
    streamRead.Open
    streamRead.LoadFromFile a_sFrom
    
    '// 改行コードLFをCRLFに変換
    sText = streamRead.ReadText
    sText = Replace(sText, vbLf, vbCrLf)
    sText = Replace(sText, vbCr & vbCr, vbCr)
    
    '// ファイル書き込み
    streamWrite.Type = 2 'adTypeText
    streamWrite.Charset = "Shift-JIS"
    streamWrite.Open
    
    '// データ書き込み
    streamWrite.WriteText sText
    
    fnm = a_sFrom
    If InStr(fnm, "\") > 0 Then
        fnm = Right(fnm, Len(fnm) - InStrRev(fnm, "\"))
    End If
    
    '// 保存
    streamWrite.SaveToFile a_sTo & fnm, 2 'adSaveCreateOverWrite
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub


Private Sub MySjisToUtf8(a_sFrom, a_sTo)
    Dim sText                           '// ファイルデータ
    Set streamRead = CreateObject("ADODB.Stream")
    Set streamWrite = CreateObject("ADODB.Stream")
    
    '// ファイル読み込み
    streamRead.Type = 2 'adTypeText
    streamRead.Charset = "Shift_JIS"
    'streamRead.LineSeparator = adCRLF
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// 改行コードCRLFをLFに変換
    sText = streamRead.ReadText
    sText = Replace(sText, vbCrLf, vbLf)
    
    '// ファイル書き込み
    streamWrite.Type = 2 'adTypeText
    streamWrite.Charset = "UTF-8"
    'streamWrite.LineSeparator = adLF
    streamWrite.Open
    '// データ書き込み
    streamWrite.WriteText sText
    
    streamWrite.position = 0
    streamWrite.Type = 1 'adTypeBinary
    streamWrite.position = 3
    Dim byteData() As Byte
    byteData = streamWrite.Read
    streamWrite.Close '一旦ストリームを閉じる（リセット）
    streamWrite.Open 'ストリームを開く
    streamWrite.Write byteData
    
    fnm = a_sFrom
    If InStr(fnm, "\") > 0 Then
        fnm = Right(fnm, Len(fnm) - InStrRev(fnm, "\"))
    End If
    
    '// 保存
    streamWrite.SaveToFile a_sTo & fnm, 2 'adSaveCreateOverWrite
    
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub


Private Sub test_sjisutf()

    MySjisToUtf8 "C:\Users\1st\Desktop\UtilFile.bas", "C:\Users\1st\Desktop\utfUtilFile.txt"
    MyUtf8ToSjis "C:\Users\1st\Desktop\utfUtilFile.txt", "C:\Users\1st\Desktop\sjisUtilFile.txt"

End Sub



Private Function MyCreateTempFolder()
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = fso.GetSpecialFolder(2) & "\" & GetGUID

    MkDir tmp
        
    MyCreateTempFolder = tmp & "\"
End Function

Private Sub MyDeleteTempFolder(tmp)
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmpbase = fso.GetSpecialFolder(2) & "\"
    If Left(tmp, Len(tmpbase)) = tmpbase Then
        If RmDirBestEffort(tmp, dummymsg) Then
            'Pass
        End If
    Else
        Debug.Print tmp & "は、" & tmpbase & "以下のフォルダではありませんので削除しません"
    End If
    
End Sub

Private Sub test_temp()

    Debug.Print MyCreateTempFolder

End Sub


 
Private Function MyCreateGuidString()
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    
    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            MyCreateGuidString = strGuid
        End If
    End If
End Function
 
Private Function MyGetGUID()
    Dim strGuid As String
    strGuid = MyCreateGuidString()
    
    strGuid = Replace(Replace(strGuid, "{", ""), "}", "")
    
    '謎の\0が末尾につくので、削除する
    MyGetGUID = Left(strGuid, Len(strGuid) - 1)
End Function



Private Function MyRmDirBestEffort(ByVal sDir As String, _
                ByRef sMsg As String, _
                Optional ByVal isOnlyFile As Boolean = False) _
                As Boolean
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder As Object
    sMsg = ""
    If Not objFSO.FolderExists(sDir) Then
        sMsg = "指定のフォルダは存在しません。"
        RmDirBestEffort = False
        Exit Function
    End If
    Set objFolder = objFSO.GetFolder(sDir)
    RmDirBestEffortRecur objFolder, isOnlyFile, sMsg
    If sMsg = "" Then
        MyRmDirBestEffort = True
    Else
        MyRmDirBestEffort = False
    End If
End Function

Private Sub RmDirBestEffortRecur(ByVal objFolder As Object, _
                    ByVal isOnlyFile As Boolean, _
                    ByRef sMsg As String)
    Dim objFolderSub As Object
    Dim objFile As Object
    On Error Resume Next
    For Each objFolderSub In objFolder.SubFolders
        Call RmDirBestEffortRecur(objFolderSub, isOnlyFile, sMsg)
    Next
    For Each objFile In objFolder.Files
        objFile.Delete
        If Err.Number <> 0 Then
            sMsg = sMsg & "ファイル「" & objFile.path & "」が削除できませんでした" & vbLf
            Err.Clear
        End If
    Next
    If Not isOnlyFile Then
        objFolder.Delete
        If Err.Number <> 0 Then
            sMsg = sMsg & "フォルダ「" & objFolder.path & "」が削除できませんでした" & vbLf
            Err.Clear
        End If
    End If
    Set objFolderSub = Nothing
    Set objFile = Nothing
    On Error GoTo 0
End Sub


'-------------------------------------------------------------------------------
'ここまで、UtilFileにもコピーした


Public Sub MyImportTask()
    Set fso = CreateObject("Scripting.FileSystemObject")
    syncfile = fso.GetSpecialFolder(2) & "\moduleimporter.txt"
    If Dir(syncfile) = "" Then
        'フォルダ選択ダイアログ（デフォルトはダウンロードフォルダ）
        
        dlpath = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.path
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = fso.GetFolder(dlpath) & "\"
            .Show
            
            If .SelectedItems.Count = 0 Then End
            sPath = .SelectedItems(1)
        End With
    Else
        'テンポラリフォルダにmoduleimporter.txtがあれば、その中に書かれたフォルダを採用
        With fso.GetFile(syncfile).OpenAsTextStream
            sPath = Trim(Replace(Replace(.ReadAll, vbCr, ""), vbLf, ""))
            .Close
        End With
    End If
    
    Debug.Print sPath
    'モジュール一括インポート
    ImportAll sPath

End Sub



Public Sub MyExportTask()
    dstpath = "D:\Dropbox (個人用)\★個人PJ\●vbasuper\VBALIB\"
    If Dir(dstpath) = "" Then End

    'マ☆テンポラリフォルダ生成
    temp = MyCreateTempFolder
    
    'マ☆モジュール一括エクスポート
    ExportAll temp
    
    'マ☆文字コード変換to UTF8
    ReDim sArModule(0)
    searchAllFile temp, sArModule
    
    'マ☆決まったフォルダにコピー
    For i = LBound(sArModule) To UBound(sArModule)
        MySjisToUtf8 sArModule(i), dstpath
        Debug.Print sArModule(i)
    Next

End Sub




Private Sub test_import()

    'ImportAll "C:\Users\1st\Desktop\test"

    'テンポラリフォルダを生成する
    'ExportAll "C:\Users\1st\Desktop\test"

End Sub



Public Sub ImportAll(a_sModulePath)
    On Error Resume Next
    
    Set oFso = CreateObject("Scripting.FileSystemObject")            '// FileSystemObjectオブジェクト
    Dim sArModule()                     '// モジュールファイル配列
    Dim sModule                                 '// モジュールファイル
    Dim sExt        As String                   '// 拡張子
    Dim iMsg                                    '// MsgBox関数戻り値
    
    pref = "Util"
    Set app = Application
    Select Case Application.Name
    Case "Microsoft Excel"
        Set ThisCode = app.ThisWorkbook
        pref = "XlsUtil"
    Case "Microsoft Word"
        Set ThisCode = ThisDocument
        pref = "DocUtil"
    Case "Microsoft PowerPoint"
        Set ThisCode = app.ActivePresentation
        pref = "PptUtil"
    Case "Outlook"
        Stop
    Case Else
        Stop
    End Select
    
    ReDim sArModule(0)
    
    '// 全モジュールのファイルパスを取得
    Call searchAllFile(a_sModulePath, sArModule)
    
    '// 全モジュールをループ
    For Each sModule In sArModule
        '// 拡張子を小文字で取得
        sExt = LCase(oFso.GetExtensionName(sModule))
        temp = sModule
        If InStr(sModule, "\") > 0 Then
            temp = Right(temp, Len(temp) - InStrRev(temp, "\"))
        End If
        tgtflag = False
        If Left(temp, Len("Util")) = "Util" Then tgtflag = True
        If Left(temp, Len(pref)) = pref Then tgtflag = True
        
        '// 拡張子がcls、frm、basのいずれかの場合
        If tgtflag And (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// 同名モジュールを削除
            Call ThisCode.VBProject.VBComponents.Remove(a_TargetBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// モジュールを追加
            Call ThisCode.VBProject.VBComponents.Import(sModule)
            '// Import確認用ログ出力
            Debug.Print sModule
        End If
    Next
End Sub


'// 指定フォルダ配下のファイルパスを取得
'// 引数１：フォルダパス
'// 引数２：ファイルパス配列
Private Sub searchAllFile(a_sFolder, s_ArFile())
    Set oFso = CreateObject("Scripting.FileSystemObject")
    Dim oFolder     As Object
    Dim oSubFolder  As Object
    Dim oFile       As Object
    Dim i
    
    '// フォルダがない場合
    If (oFso.FolderExists(a_sFolder) = False) Then
        Exit Sub
    End If
    
    Set oFolder = oFso.GetFolder(a_sFolder)
    
    '// サブフォルダを再帰（サブフォルダを探す必要がない場合はこのFor文を削除してください）
    For Each oSubFolder In oFolder.SubFolders
        Call searchAllFile(oSubFolder.path, s_ArFile)
    Next
    
    i = UBound(s_ArFile)
    
    '// カレントフォルダ内のファイルを取得
    For Each oFile In oFolder.Files
        If (i <> 0 Or s_ArFile(i) <> "") Then
            i = i + 1
            ReDim Preserve s_ArFile(i)
        End If
        
        '// ファイルパスを配列に格納
        s_ArFile(i) = oFile.path
    Next
End Sub


Public Sub ExportAll(sPath)
    Dim module                  As Object      '// モジュール
    Dim moduleList              As Object     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook                                  '// 処理対象ブックオブジェクト
    
    sOutPath = sPath
    If Right(sPath, 1) <> "\" Then sOutPath = sOutPath & "\"
    
    pref = "Util"
    Set app = Application
    Select Case Application.Name
    Case "Microsoft Excel"
        Set ThisCode = app.ThisWorkbook
        pref = "XlsUtil"
    Case "Microsoft Word"
        Set ThisCode = app.ThisDocument
        pref = "DocUtil"
    Case "Microsoft PowerPoint"
        Set ThisCode = app.ActivePresentation
        pref = "PptUtil"
    Case "Outlook"
        Stop
    Case Else
        Stop
    End Select
    
    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = ThisCode.VBProject.VBComponents
    
    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        skipflag = True
        If Left(module.Name, Len("Util")) = "Util" Then skipflag = False
        If Left(module.Name, Len(pref)) = pref Then skipflag = False
        'If module.Name = "PortMacro" Then skipflag = False
        If skipflag Then GoTo CONTINUE
    
        '// クラス
        If (module.Type = 2) Then 'vbext_ct_ClassModule
            extension = "cls"
        '// フォーム
        ElseIf (module.Type = 3) Then 'vbext_ct_MSForm
            '// .frxも一緒にエクスポートされる
            extension = "frm"
        '// 標準モジュール
        ElseIf (module.Type = 1) Then 'vbext_ct_StdModule
            extension = "bas"
        '// その他
        Else
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        '// エクスポート実施
        sFilePath = sOutPath & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        '// 出力先確認用ログ出力
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

