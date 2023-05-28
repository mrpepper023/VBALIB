Attribute VB_Name = "UtilFile"
'=========================================================================================
'UtilFile 20230527
'
'UtilFile�͎�Ƀt�@�C������������AExcel VBA�Ɉˑ����Ȃ��R�[�h���W�߂�����
'=========================================================================================
'�t�@�C�����X�g�𓾂�
'Public Function GetFileList(ByVal path, Optional ext = "")
'�t�@�C�����X�g�𓾂�i�t�@�C�����𐳋K�\���Ŏw�肷��j
'Public Function GetFileListRegex(ByVal path, Optional ByVal recur = False, Optional pat = ".*")
'�t�H���_�I���_�C�A���O
'Public Function FolderPicker(defpath)
'�t�@�C�����ɗǂ��d���ގ���������
'Public Function TimeString()
'=========================================================================================


' �t�@�C���p�X�w�� -----------------------------------------------------------

'�t�@�C���ꗗ���ċA�I�Ɏ擾����
Public Function GetFileList(ByVal path, Optional ext = "")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arr = CreateObject("System.Collections.ArrayList")

    GetFileListRecur arr, fso, fso.GetFolder(path), LCase(ext)
    
    Set fso = Nothing
    GetFileList = arr.ToArray
    Set arr = Nothing

End Function

'�t�@�C���ꗗ�擾�̍ċA�P��
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

'�������������K�\���ł��ق�����
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

'�t�@�C���ꗗ�擾
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



'�p�X�ϊ����s���֐��ł��������́BSharePoint�ɂ��̎�@�ŃA�N�Z�X�ł��Ȃ��Ȃ��Ă���̂ŁA�P�Ȃ郔�@���f�[�^
Public Function cnvNetPath2Local(path As String) As String
    
    If InStr(path, "http://") > 0 Then
        MsgBox "���[�J���t�H���_�Ŏ��s���Ă�������"
        End
    ElseIf InStr(path, "https://") > 0 Then
        MsgBox "���[�J���t�H���_�Ŏ��s���Ă�������"
        End
    Else
        ' ���o�ł��Ȃ��ꍇ�͂��̂܂܃p�X��Ԃ�
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

    basepath = ThisWorkbook.Sheets("���s").Range("B4")
    If basepath = "" Then basepath = ThisWorkbook.path
    basepath = FolderPicker(basepath)
    If basepath = "" Then End
    ThisWorkbook.Sheets("���s").Range("B4") = basepath

End Sub



Public Function TimeString()

    TimeString = Format(Now(), "yyyymmdd_hhnnss")

End Function






