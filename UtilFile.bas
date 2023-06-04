Attribute VB_Name = "UtilFile"
'=========================================================================================
'UtilFile 20230527
'
'UtilFileは主にファイル操作を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'UTF8のファイルをSJISに変換する
'Sub Utf8ToSjis(a_sFrom, a_sTo)
'SJISのファイルをUTF8に変換する
'Sub SjisToUtf8(a_sFrom, a_sTo)
'テンポラリフォルダを作成してフォルダの「\」つきフルパスを返す
'Public Function CreateTempFolder()
'テンポラリフォルダを削除する（テンポラリフォルダでなければ消さないので安全
'Public Sub DeleteTempFolder(tmp)
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
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr


Sub Utf8ToSjis(a_sFrom, a_sTo)
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
    
    '// 保存
    streamWrite.SaveToFile a_sTo, 2 'adSaveCreateOverWrite
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub


Sub SjisToUtf8(a_sFrom, a_sTo)
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
    
    streamWrite.Position = 0
    streamWrite.Type = 1 'adTypeBinary
    streamWrite.Position = 3
    Dim byteData() As Byte
    byteData = streamWrite.Read
    streamWrite.Close '一旦ストリームを閉じる（リセット）
    streamWrite.Open 'ストリームを開く
    streamWrite.Write byteData
    
    '// 保存
    streamWrite.SaveToFile a_sTo, 2 'adSaveCreateOverWrite
    
    '// クローズ
    streamRead.Close
    streamWrite.Close
End Sub


Private Sub test_sjisutf()

    SjisToUtf8 "C:\Users\1st\Desktop\UtilFile.bas", "C:\Users\1st\Desktop\utfUtilFile.txt"
    Utf8ToSjis "C:\Users\1st\Desktop\utfUtilFile.txt", "C:\Users\1st\Desktop\sjisUtilFile.txt"

End Sub



Public Function CreateTempFolder()
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = fso.GetSpecialFolder(2) & "\" & GetGUID

    MkDir tmp
        
    CreateTempFolder = tmp & "\"
End Function

Public Sub DeleteTempFolder(tmp)
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

    Debug.Print CreateTempFolder

End Sub


 
Private Function CreateGuidString()
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
            CreateGuidString = strGuid
        End If
    End If
End Function
 
Public Function GetGUID()
    Dim strGuid As String
    strGuid = CreateGuidString()
    
    strGuid = Replace(Replace(strGuid, "{", ""), "}", "")
    
    '謎の\0が末尾につくので、削除する
    GetGUID = Left(strGuid, Len(strGuid) - 1)
End Function



Function RmDirBestEffort(ByVal sDir As String, _
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
        RmDirBestEffort = True
    Else
        RmDirBestEffort = False
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



' ファイルパス指定 -----------------------------------------------------------

'ファイル一覧を再帰的に取得する
Public Function GetFileList(ByVal path, Optional ext = "")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arr = CreateObject("System.Collections.ArrayList")

    GetFileListRecur arr, fso, fso.GetFolder(path), LCase(ext)
    
    Set fso = Nothing
    GetFileList = arr.ToArray
    Set arr = Nothing

End Function

'ファイル一覧取得の再帰単位
Private Sub GetFileListRecur(ByRef arr, ByRef fso, ByVal folder, ByRef ext)
    
    For Each file In folder.Files
        fileext = fso.GetExtensionName(file.path)
        If Len(ext) = 0 Or ext = Left(LCase(fileext), Len(ext)) Then
            arr.Add file.path
        End If
    Next
  
    For Each D In folder.SubFolders
        GetFileListRecur arr, fso, D, ext
    Next

End Sub

'★★いっそ正規表現版もほしいね
Public Function GetFileListRegex(ByVal path, Optional ByVal recur = False, Optional pat = ".*")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arr = CreateObject("System.Collections.ArrayList")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.ignorecase = True
    re.Pattern = pat

    GetFileListRegexRecur arr, fso, fso.GetFolder(path), recur, re
    
    Set fso = Nothing
    GetFileListRegex = arr.ToArray
    Set arr = Nothing

End Function

'ファイル一覧取得
Private Sub GetFileListRegexRecur(ByRef arr, ByRef fso, ByVal folder, ByVal recur, ByRef re)
    
    For Each file In folder.Files
        fname = fso.GetFileName(file.path)
        If re.test(vbCr & fname & vbCr) Then
            arr.Add file.path
        End If
    Next
  
    If recur Then
        For Each D In folder.SubFolders
            GetFileListRegexRecur arr, fso, D, recur, re
        Next
    End If

End Sub



'パス変換を行う関数であったもの。SharePointにこの手法でアクセスできなくなっているので、単なるヴァリデータ
Public Function cnvNetPath2Local(path As String) As String
    
    If InStr(path, "http://") > 0 Then
        MsgBox "ローカルフォルダで実行してください"
        End
    ElseIf InStr(path, "https://") > 0 Then
        MsgBox "ローカルフォルダで実行してください"
        End
    Else
        ' 検出できない場合はそのままパスを返す
        cnvNetPath2Local = path
    End If

End Function





Public Function FolderPicker(defpath)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = CreateObject("Scripting.FileSystemObject").GetFolder(defpath) & "\"
        .Show
        
        If .SelectedItems.Count = 0 Then
            FolderPicker = ""
            Exit Function
        End If
        FolderPicker = cnvNetPath2Local(.SelectedItems(1))
    End With
End Function

Private Sub test_FolderPicker2()

    If basepath = "" Then basepath = ThisWorkbook.path
    basepath = FolderPicker(basepath)
    If basepath = "" Then End

End Sub

Private Sub test_FolderPicker()

    basepath = ThisWorkbook.Sheets("実行").Range("B4")
    If basepath = "" Then basepath = ThisWorkbook.path
    basepath = FolderPicker(basepath)
    If basepath = "" Then End
    ThisWorkbook.Sheets("実行").Range("B4") = basepath

End Sub



Public Function TimeString()

    TimeString = Format(Now(), "yyyymmdd_hhnnss")

End Function






