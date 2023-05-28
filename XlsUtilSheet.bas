Attribute VB_Name = "XlsUtilSheet"
Const DANGERFAST = True
'=========================================================================================
'XlsUtilSheet 20230527
'
'XlsUtilSheet��
'=========================================================================================
'���[�N�V�[�g�����ɁA���p�ϊ�����֐�
'Public Function HanKana(str)
'�V�[�g�I�u�W�F�N�g�Ɨ�ԍ���n���A�Y����̍ŏI�s�̍s�ԍ����擾����
'Public Function GetMaxRow(ByRef sh, Optional ByVal col = 1)
'�V�[�g�I�u�W�F�N�g�ƍs�ԍ���n���A�Y���s�̍ŏI��̗�ԍ����擾����
'Public Function GetMaxCol(ByRef sh, Optional ByVal row = 1)
'�V�[�g�I�u�W�F�N�g��n���A�g�p�����ŏI�s�̍s�ԍ����擾����
'Public Function GetUsedMaxRow(ByRef sh)
'�V�[�g�I�u�W�F�N�g��n���A�g�p�����ŏI��̗�ԍ����擾����
'Public Function GetUsedMaxCol(ByRef sh)
'�V�[�g�I�u�W�F�N�g�ƍ���E���̒��l�����ɂ���Range�I�u�W�F�N�g��Ԃ�
'Public Function RectRange(ByRef sheet, ByVal r_top, ByVal c_left, Optional ByVal r_bottom = 0, Optional ByVal c_right = 0)
'Range�I�u�W�F�N�g��n���āA��ҏW�Z���̑������s��
'Public Sub DecorateManualCells(ByRef rng, Optional ByVal defval = "")
'Excel�}�N���̒P���ȍ������F�G���[�������ʓ|�ɂȂ�̂ŁA�o�O���Ƃ�Ă���d�������B
'Public Sub FastSetting(flag)
'�r�W�[���[�v������Ȃ�A���[�v���ł�����Ă�ł���
'Public Sub FastSettingDoEvents(Optional str = "")
'�ǂ����Ă��~�߂����ꍇ�A������g��
'Public Sub FastSettingStop(Optional str = "")
'�R�����g��Statusbar��Immediate Window�ɕ\��
'Public Sub d_print_____(str)
'A,B,AA,AB�Ȃǂ̗�\�L�̃A���t�@�x�b�g���A1���珇�ɐU������ԍ��ɕϊ�����
'Public Function addr2col(ByVal str)
'1���珇�ɐU������ԍ����AA,B,AA,AB�Ȃǂ̗�\�L�̃A���t�@�x�b�g�ɕϊ�����
'Public Function col2addr(ByVal num)
'targetstr�̕�����̖`�����Aprefix�ɂȂ��Ă��邩�ǂ����𔻒肷��
'Public Function StartsWith(ByRef targetstr, ByRef prefix)
'targetstr�̕�����̖������Asuffix�ɂȂ��Ă��邩�ǂ����𔻒肷��
'Public Function EndsWith(ByRef targetstr, ByRef suffix)
'targetstr�̐擪����Aprefix���폜����B�擪����v���Ă��Ȃ���΁Atargetstr�����̂܂ܕԂ�
'Public Function RemovePrefix(ByRef targetstr, ByRef prefix)
'targetstr�̖�������Asuffix���폜����B��������v���Ă��Ȃ���΁Atargetstr�����̂܂ܕԂ�
'Public Function RemoveSuffix(ByRef targetstr, ByRef suffix)
'Excel��p�I�w�肵���t�@�C�����J����Ă���΂��̃u�b�N���A�����Ȃ��΃t�@�C�����J���ău�b�N�𓾂�
'Public Function WiseOpen(path, ByRef closeflag)
'�����Ŏ��������O�̃V�[�g��T���āA���݂���Ȃ�TRUE�A���݂��Ȃ��Ȃ�FALSE��Ԃ�
'Public Function FindSheet(ByVal str)
'�����Ŏ��������O�̃V�[�g��T���āA���݂��Ȃ��Ȃ�G���[���b�Z�[�W��\�����Ē��f����
'Public Sub FindSheet_Trap(ByVal str)
'����FindSheetRegex������Ă�������
'pref�Ŏw�肵��������Ŏn�܂閼�O�̃V�[�g���A�z��ϐ��Ɋi�[���ĕԂ�
'Public Function FindSheetPrefix(ByVal pref)
'�V�[�g�ւ̎Q�Ƃƕ�����i�ƍs�ԍ��j���w�肵�āA����̕�����ƈ�v����\��̂�����������ĕԂ��i������Ȃ��ꍇ��0��Ԃ�
'Public Function FindCol(ByRef sheet, ByRef str, Optional ByVal row = 1)
'�����V�[�g�̕\��s�̐���������Ă�������
'����������擾�Ƃ��A�����Ă������B
'=========================================================================================




'���[�N�V�[�g�����ɁA���p�ϊ�����֐�
Public Function HanKana(str)

    HanKana = StrConv(str, vbNarrow)
    
End Function

'�V�[�g�I�u�W�F�N�g�Ɨ�ԍ���n���A�Y����̍ŏI�s�̍s�ԍ����擾����
Public Function GetMaxRow(ByRef sh, Optional ByVal col = 1)
    
    GetMaxRow = sh.Cells(sh.Cells(sh.Rows.Count, col).row, col).End(xlUp).row
    
End Function

'�V�[�g�I�u�W�F�N�g�ƍs�ԍ���n���A�Y���s�̍ŏI��̗�ԍ����擾����
Public Function GetMaxCol(ByRef sh, Optional ByVal row = 1)
    
    GetMaxCol = sh.Cells(row, sh.Cells(row, sh.Columns.Count).col).End(xlToLeft).Column
    
End Function '�V�[�g�I�u�W�F�N�g��n���A�g�p�����ŏI�s�̍s�ԍ����擾����

'�V�[�g�I�u�W�F�N�g��n���A�g�p�����ŏI�s�̍s�ԍ����擾����
Public Function GetUsedMaxRow(ByRef sh)
    
    GetUsedMaxRow = sh.UsedRange.Rows(sh.UsedRange.Rows.Count).row
    
End Function

'�V�[�g�I�u�W�F�N�g��n���A�g�p�����ŏI��̗�ԍ����擾����
Public Function GetUsedMaxCol(ByRef sh)
    
    GetUsedMaxCol = sh.UsedRange.Columns(sh.UsedRange.Columns.Count).Column
    
End Function



'�V�[�g�I�u�W�F�N�g�ƍ���E���̒��l�����ɂ���Range�I�u�W�F�N�g��Ԃ�
Public Function RectRange(ByRef sheet, ByVal r_top, ByVal c_left, Optional ByVal r_bottom = 0, Optional ByVal c_right = 0)

    If r_bottom = 0 Then r_bottom = r_top
    If c_right = 0 Then c_right = c_left
    With sheet
        Set RectRange = .Range(.Cells(r_top, c_left), .Cells(r_bottom, c_right))
    End With
    
End Function



'Range�I�u�W�F�N�g��n���āA��ҏW�Z���̑������s��
Public Sub DecorateManualCells(ByRef rng, Optional ByVal defval = "")
    
    If Len(defval) > 0 Then
        rng.Value = defval
    End If
    
    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .Weight = xlThin
    End With
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With rng.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 255, 0)
    End With
    
End Sub





'Excel�}�N���̒P���ȍ������F�G���[�������ʓ|�ɂȂ�̂ŁA�o�O���Ƃ�Ă���d�������B
Public Sub FastSetting(flag)
    If DANGERFAST Then
        If flag Then
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
        Else
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Application.Calculation = xlCalculationAutomatic
        End If
    End If
End Sub

'�r�W�[���[�v������Ȃ�A���[�v���ł�����Ă�ł���
Public Sub FastSettingDoEvents(Optional str = "")
    
    If Rnd < 0.01 Then
        savedsu = Application.ScreenUpdating
        savedee = Application.EnableEvents
        savedcc = Application.Calculation
        
        FastSetting False
        
        Application.StatusBar = str
        Debug.Print "[" & Format(Time(), "yyyy/mm/dd hh:nn:ss") & "] " & str
        DoEvents
        
        Application.ScreenUpdating = savedsu
        Application.EnableEvents = savedee
        Application.Calculation = savedcc
    End If
    
End Sub

'�ǂ����Ă��~�߂����ꍇ�A������g��
Public Sub FastSettingStop(Optional str = "")
    
    If Rnd < 0.01 Then
        savedsu = Application.ScreenUpdating
        savedee = Application.EnableEvents
        savedcc = Application.Calculation
        
        FastSetting False
        
        Application.StatusBar = str
        Debug.Print "[" & Format(Time(), "yyyy/mm/dd hh:nn:ss") & "] " & str
        Stop
        
        Application.ScreenUpdating = savedsu
        Application.EnableEvents = savedee
        Application.Calculation = savedcc
    End If
    
End Sub





' �R�����g��Statusbar��Immediate Window�ɕ\��
Public Sub d_print_____(str)
    savedsu = Application.ScreenUpdating
    savedee = Application.EnableEvents
    savedcc = Application.Calculation
    FastSetting False

    Application.StatusBar = str
    Debug.Print "[" & Format(Time(), "yyyy/mm/dd hh:nn:ss") & "] " & str
    DoEvents
    
    Application.ScreenUpdating = savedsu
    Application.EnableEvents = savedee
    Application.Calculation = savedcc
End Sub



' A,B,AA,AB�Ȃǂ̗�\�L�̃A���t�@�x�b�g���A1���珇�ɐU������ԍ��ɕϊ�����
Public Function addr2col(ByVal str)
    On Error GoTo e
    addr2col = ActiveSheet.Range(str & 1).Column
    On Error GoTo 0
    Exit Function
e:
    On Error GoTo 0
    addr2col = False
End Function

'1���珇�ɐU������ԍ����AA,B,AA,AB�Ȃǂ̗�\�L�̃A���t�@�x�b�g�ɕϊ�����
Public Function col2addr(ByVal num)
    On Error GoTo e
    col2addr = Split(ActiveSheet.Cells(1, num).Address(True, False), "$")(0)
    On Error GoTo 0
    Exit Function
e:
    On Error GoTo 0
    col2addr = False
End Function



Private Sub test_addr2col()
    Debug.Print addr2col("A")
    Debug.Print addr2col("AA")
    Debug.Print addr2col("Xfd")
    Debug.Print addr2col("XXX")
    Debug.Print col2addr(2)
    Debug.Print col2addr(23)
    Debug.Print col2addr(235)
    Debug.Print col2addr(16251)
    Debug.Print col2addr(46251)
End Sub



'targetstr�̕�����̖`�����Aprefix�ɂȂ��Ă��邩�ǂ����𔻒肷��
' �u���l�}�s�v�u���l�v�Ȃ�TRUE�u�ē���H�v�u���H�v�Ȃ�FALSE
Public Function StartsWith(ByRef targetstr, ByRef prefix)
    StartsWith = (Left(targetstr, Len(prefix)) = prefix)
End Function
'targetstr�̕�����̖������Asuffix�ɂȂ��Ă��邩�ǂ����𔻒肷��
Public Function EndsWith(ByRef targetstr, ByRef suffix)
    EndsWith = (Right(targetstr, Len(suffix)) = suffix)
End Function

'targetstr�̐擪����Aprefix���폜����B�擪����v���Ă��Ȃ���΁Atargetstr�����̂܂ܕԂ�
' �u���l�}�s�v�u���l�v�Ȃ�Ԓl�́u�}�s�v
Public Function RemovePrefix(ByRef targetstr, ByRef prefix)
    If StartsWith(targetstr, prefix) Then
        RemovePrefix = Right(targetstr, Len(targetstr) - Len(prefix))
        Exit Function
    End If
    RemovePrefix = targetstr
End Function
' targetstr�̖�������Asuffix���폜����B��������v���Ă��Ȃ���΁Atargetstr�����̂܂ܕԂ�
Public Function RemoveSuffix(ByRef targetstr, ByRef suffix)
    If EndsWith(targetstr, suffix) Then
        RemoveSuffix = Left(targetstr, Len(targetstr) - Len(suffix))
        Exit Function
    End If
    RemoveSuffix = targetstr
End Function



'�����t�@�C���I�[�v���i���ɊJ���Ă���΂�������Q�Ɓj���ău�b�N�I�u�W�F�N�g��Ԃ�
'�g�p��Acloseflag���Q�Ƃ���True�Ȃ�t�@�C�������̂𐄏�����
'�Ȃ��AWiseOpen�͊�{�I�ɓǂݎ���p�ŊJ���i�����Ώۂ����������d�l�őI�Ԃ̂͂��Ԃˁ[�j
'�Ȃ��AExcel��p
Public Function WiseOpen(path, ByRef closeflag)
    Set fso = CreateObject("scripting.filesystemobject")
    closeflag = False
    
    fname = fso.GetFileName(path)
    For Each f In Workbooks
        If fname = f.Name Then
            Set WiseOpen = f
            GoTo exitwiseopen
        End If
    Next
    
    closeflag = True
    savedupdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set WiseOpen = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, CorruptLoad:=xlRepairFile)
    Application.Windows(fname).Visible = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = savedupdating
    
exitwiseopen:
    Set fso = Nothing
End Function

Private Sub testwiseopen()

    Set bk = WiseOpen("D:\Users\miyokomizo\Desktop\�ҏW�F���������d����L���E�ؕ��i2021�N3��1���`2022�N2��28���j.xlsx", closeflag)
    If bk Is Nothing Then End
    
    If closeflag Then bk.Close savechanges:=False
    Set bk = Nothing

End Sub





' �����Ŏ��������O�̃V�[�g��T���āA���݂���Ȃ�TRUE�A���݂��Ȃ��Ȃ�FALSE��Ԃ�
Public Function FindSheet(ByVal str)
    For Each N In ThisWorkbook.Sheets
        If N.Name = str Then
            FindSheet = True
            Exit Function
        End If
    Next
    
    FindSheet = False
End Function

' �����Ŏ��������O�̃V�[�g��T���āA���݂��Ȃ��Ȃ�G���[���b�Z�[�W��\�����Ē��f����
Public Sub FindSheet_Trap(ByVal str)

    If Not FindSheet(str) Then
        MsgBox str & "��������܂���"
        End
    End If

End Sub

Private Sub test_findsheet()

    FindSheet_Trap "RAW1"
    FindSheet_Trap "RAW2"
    FindSheet_Trap "���H�σf�[�^�W�v�ݒ�"
    FindSheet_Trap "RAW1������"
    FindSheet_Trap "RAW2������"

End Sub



' pref�Ŏw�肵��������Ŏn�܂閼�O�̃V�[�g���A�z��ϐ��Ɋi�[���ĕԂ�
' �񋓂���ۂɂ�lbound,ubound���g������
Public Function FindSheetPrefix(ByVal pref)
    Set temp = CreateObject("system.collections.arraylist")
    
    preflen = Len(pref)
    For Each N In ThisWorkbook.Sheets
        If Left(N.Name, preflen) = pref Then
            temp.Add N.Name
        End If
    Next
    
    FindSheetPrefix = temp.ToArray()
    Set temp = Nothing

End Function

Private Sub test_FindSheetPrefix()

    a = FindSheetPrefix("RAW")
    For i = LBound(a) To UBound(a)
        Debug.Print a(i)
    Next

End Sub



' �V�[�g�ւ̎Q�Ƃƕ�����i�ƍs�ԍ��j���w�肵�āA����̕�����ƈ�v����\��̂�����������ĕԂ�
' ������Ȃ��ꍇ��0��Ԃ�
Public Function FindCol(ByRef sheet, ByRef str, Optional ByVal row = 1)
    maxc = MaxColumn(sheet, row)
    For i = 1 To maxc
        If str = sheet.Cells(row, i) Then
            FindCol = i
            Exit Function
        End If
    Next
    FindCol = 0
End Function

Private Sub test_FindCol()

    Debug.Print FindCol(ThisWorkbook.Sheets("���H�σf�[�^�W�v�ݒ�"), "�e�X�g�ݖ�")
    Debug.Print FindCol(ThisWorkbook.Sheets("���H�σf�[�^�W�v�ݒ�"), "���݂��Ȃ���")

End Sub







