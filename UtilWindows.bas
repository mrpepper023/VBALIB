Attribute VB_Name = "UtilWindows"
'=========================================================================================
'UtilWindows 20230527
'
'UtilWindows�͎��Windows OS�̐���������AExcel VBA�Ɉˑ����Ȃ��R�[�h���W�߂�����
'=========================================================================================
'�N���b�v�{�[�h�Ƀv���[���e�L�X�g���Z�b�g����
'Public Sub SetClip(txt)
'�N���b�v�{�[�h����v���[���e�L�X�g��ǂݎ��
'Public Function GetClip()
'URL�ƃ��\�b�h���w�肵�ăE�F�u�ɃA�N�Z�X���A���ʂ𕶎���œ���
'Public Function HostApplication()
'���̃}�N�����ǂ̃A�v���ɑg�ݍ��܂�Ă��邩"Microsoft Excel"�Ƃ��ŕ��򂷂邽��
'Public Function EscapedSplit(txt, delim)
'�������z��̎������𓾂�
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
        'SetCursorPos�@�E�E�E�}�E�X�𓮂����E�}�E�X�̃|�C���^�[�̑�����s���B
        Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
        'Mouseevent�@�@�E�E�E�}�E�X���N���b�N���鑀����s���B
        Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, Optional ByVal dx As Long = 0, Optional ByVal dy As Long = 0, Optional ByVal dwDate As Long = 0, Optional ByVal dwExtraInfo As Long = 0)
        'GetCursorPos�@�E�E�E�}�E�X�̃|�C���^�[�̈ʒu���擾���܂��B
        Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As coord) As Long
    #Else
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
        'SetCursorPos�@�E�E�E�}�E�X�𓮂����E�}�E�X�̃|�C���^�[�̑�����s���B
        Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
        'Mouseevent�@�@�E�E�E�}�E�X���N���b�N���鑀����s���B
        Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, Optional ByVal dx As Long = 0, Optional ByVal dy As Long = 0, Optional ByVal dwDate As Long = 0, Optional ByVal dwExtraInfo As Long = 0)
        'GetCursorPos�@�E�E�E�}�E�X�̃|�C���^�[�̈ʒu���擾���܂��B
        Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As coord) As Long
    #End If
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If



'�d�v�IExcel�ȊO�ł����ʃR�[�h�Ŏ��s���邽�߂ɂ́A���L�̂悤�ɂ��ׂ�
Private Sub test_multi_host()

    '�d�v�IExcel�ȊO�Ŏ��s�ł��Ȃ��R�[�h�̎��s��h��
    If Application.Name = "Microsoft Excel" Then
        Debug.Print ThisWorkbook.Name
        '�T�u���[�`���^�֐��P�ʂ̃R���p�C�����̃G���[��h�����߁AApplication���Q�ƌo�R�Œ@��
        Set xlapp = Application
        '�T�u���[�`���^�֐��P�ʂ̃R���p�C�����̃G���[��h�����߁AApplication���Q�ƌo�R�Œ@��
        xlapp.ActiveSheet.Range("A2") = "aaa"
        xlapp.ActiveSheet.Range("A2").Clear
        '�ȗ��L�@�������Ȃ������ŁA�������R�ɏ�����
    End If

End Sub




'�N���b�v�{�[�h����

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
'0: �e�L�X�g
'2: �摜
'9: BitMap
'47: �t�@�C���p�X
'14:�摜�n�H
'17:�摜�n�H
'22:�摜�n�H
'31:�摜�n�H
'45:�摜�n�H
    
    If flag.exists(0) Then
        GetClip = dao.GetText
    Else
        GetClip = ""
    End If
    
    'GetClip = dao.GetImage
End Function



'�E�F�u�A�N�Z�X�i�����API�@���p�j
'����Edge��DOM��`�����@������炵��

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



'�I�[�g�p�C���b�g�n

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



