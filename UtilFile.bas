Attribute VB_Name = "UtilFile"
'=========================================================================================
'UtilFile 20230527
'
'UtilFileは主にファイル操作を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'ファイルリストを得る
'Public Function GetFileList(ByVal path, Optional ext = "")
'ファイルリストを得る（ファイル名を正規表現で指定する）
'Public Function GetFileListRegex(ByVal path, Optional ByVal recur = False, Optional pat = ".*")
'フォルダ選択ダイアログ
'Public Function FolderPicker(defpath)
'ファイル名に良く仕込む時刻文字列
'Public Function TimeString()
'=========================================================================================


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






