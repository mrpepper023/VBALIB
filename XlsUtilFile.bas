Attribute VB_Name = "XlsUtilFile"
'=========================================================================================
'XlsUtilFile 20230527
'
'XlsUtilFileは主にファイル操作を扱う、Excel VBAに依存しないコードを集めたもの
'=========================================================================================
'テンポラリフォルダを作成してフォルダの「\」つきフルパスを返す
'Public Function CreateTempFolder()
'テンポラリフォルダを削除する（テンポラリフォルダでなければ消さないので安全
'Public Sub DeleteTempFolder(tmp)
'=========================================================================================


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


