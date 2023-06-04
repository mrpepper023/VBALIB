Attribute VB_Name = "UtilFile"
'=========================================================================================
'UtilFile 20230527
'
'UtilFile�͎�Ƀt�@�C������������AExcel VBA�Ɉˑ����Ȃ��R�[�h���W�߂�����
'=========================================================================================
'UTF8�̃t�@�C����SJIS�ɕϊ�����
'Sub Utf8ToSjis(a_sFrom, a_sTo)
'SJIS�̃t�@�C����UTF8�ɕϊ�����
'Sub SjisToUtf8(a_sFrom, a_sTo)
'�e���|�����t�H���_���쐬���ăt�H���_�́u\�v���t���p�X��Ԃ�
'Public Function CreateTempFolder()
'�e���|�����t�H���_���폜����i�e���|�����t�H���_�łȂ���Ώ����Ȃ��̂ň��S
'Public Sub DeleteTempFolder(tmp)
'GUID�𐶐�����
'Public Function GetGUID()
'�t�H���_�����ʂ̃t�@�C����T�u�f�B���N�g���܂߁A�\�Ȍ���폜����
'Function RmDirBestEffort(ByVal sDir As String, ByRef sMsg As String, Optional ByVal isOnlyFile As Boolean = False) As Boolean
'�t�@�C�����X�g�𓾂�
'Public Function GetFileList(ByVal path, Optional ext = "")
'�t�@�C�����X�g�𓾂�i�t�@�C�����𐳋K�\���Ŏw�肷��j
'Public Function GetFileListRegex(ByVal path, Optional ByVal recur = False, Optional pat = ".*")
'�t�H���_�I���_�C�A���O
'Public Function FolderPicker(defpath)
'�t�@�C�����ɗǂ��d���ގ���������
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
    Dim sText                           '// �t�@�C���f�[�^
    Set streamRead = CreateObject("ADODB.Stream")
    Set streamWrite = CreateObject("ADODB.Stream")
    
    '// �t�@�C���ǂݍ���
    streamRead.Type = 2 'adTypeText
    streamRead.Charset = "UTF-8"
    streamRead.Open
    streamRead.LoadFromFile a_sFrom
    
    '// ���s�R�[�hLF��CRLF�ɕϊ�
    sText = streamRead.ReadText
    sText = Replace(sText, vbLf, vbCrLf)
    sText = Replace(sText, vbCr & vbCr, vbCr)
    
    '// �t�@�C����������
    streamWrite.Type = 2 'adTypeText
    streamWrite.Charset = "Shift-JIS"
    streamWrite.Open
    
    '// �f�[�^��������
    streamWrite.WriteText sText
    
    '// �ۑ�
    streamWrite.SaveToFile a_sTo, 2 'adSaveCreateOverWrite
    
    '// �N���[�Y
    streamRead.Close
    streamWrite.Close
End Sub


Sub SjisToUtf8(a_sFrom, a_sTo)
    Dim sText                           '// �t�@�C���f�[�^
    Set streamRead = CreateObject("ADODB.Stream")
    Set streamWrite = CreateObject("ADODB.Stream")
    
    '// �t�@�C���ǂݍ���
    streamRead.Type = 2 'adTypeText
    streamRead.Charset = "Shift_JIS"
    'streamRead.LineSeparator = adCRLF
    streamRead.Open
    Call streamRead.LoadFromFile(a_sFrom)
    
    '// ���s�R�[�hCRLF��LF�ɕϊ�
    sText = streamRead.ReadText
    sText = Replace(sText, vbCrLf, vbLf)
    
    '// �t�@�C����������
    streamWrite.Type = 2 'adTypeText
    streamWrite.Charset = "UTF-8"
    'streamWrite.LineSeparator = adLF
    streamWrite.Open
    '// �f�[�^��������
    streamWrite.WriteText sText
    
    streamWrite.Position = 0
    streamWrite.Type = 1 'adTypeBinary
    streamWrite.Position = 3
    Dim byteData() As Byte
    byteData = streamWrite.Read
    streamWrite.Close '��U�X�g���[�������i���Z�b�g�j
    streamWrite.Open '�X�g���[�����J��
    streamWrite.Write byteData
    
    '// �ۑ�
    streamWrite.SaveToFile a_sTo, 2 'adSaveCreateOverWrite
    
    '// �N���[�Y
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
        Debug.Print tmp & "�́A" & tmpbase & "�ȉ��̃t�H���_�ł͂���܂���̂ō폜���܂���"
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
    
    '���\0�������ɂ��̂ŁA�폜����
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
        sMsg = "�w��̃t�H���_�͑��݂��܂���B"
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
            sMsg = sMsg & "�t�@�C���u" & objFile.path & "�v���폜�ł��܂���ł���" & vbLf
            Err.Clear
        End If
    Next
    If Not isOnlyFile Then
        objFolder.Delete
        If Err.Number <> 0 Then
            sMsg = sMsg & "�t�H���_�u" & objFolder.path & "�v���폜�ł��܂���ł���" & vbLf
            Err.Clear
        End If
    End If
    Set objFolderSub = Nothing
    Set objFile = Nothing
    On Error GoTo 0
End Sub



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






