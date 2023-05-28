Attribute VB_Name = "PptExport"




Private Sub analize_shape_backup(ByRef shape, ByRef dictofdict)
    Debug.Print TypeName(shape)

    Debug.Print "----------------------------------"
'    Debug.Print "Adjustments: " & shape.Adjustments '
    Debug.Print "AlternativeText: " & shape.AlternativeText '
    Debug.Print "AutoShapeType: " & shape.AutoShapeType '
    Debug.Print "BackgroundStyle: " & shape.BackgroundStyle '
    Debug.Print "BlackWhiteMode: " & shape.BlackWhiteMode '

'    Debug.Print "BottomRightCell: " & shape.BottomRightCell ''''for excel
'    Debug.Print "TopLeftCell: " & shape.TopLeftCell '
'    Debug.Print "Placement: " & shape.Placement '�I�u�W�F�N�g�Ƃ��̉��ɂ���Z���Ƃ̈ʒu�֌W��\�� XlPlacement

'    Debug.Print "Callout: " & shape.Callout '�����o���̏����v���p�e�B
    Debug.Print "Child: " & shape.Child '
'    Debug.Print "Parent: " & shape.Parent '
'    Debug.Print "ParentGroup: " & shape.ParentGroup '
    
'    Debug.Print "ControlFormat: " & shape.ControlFormat 'for excel?
    Debug.Print "Application: " & shape.Application '�w�肳�ꂽ�I�u�W�F�N�g���쐬���� Application �I�u�W�F�N�g��Ԃ��܂��B
    Debug.Print "Creator: " & shape.Creator '�I�u�W�F�N�g���쐬���ꂽ�A�v���P�[�V���������� 32 �r�b�g�̐���
'    Debug.Print "LinkFormat: " & shape.LinkFormat '�����N���ꂽ OLE �I�u�W�F�N�g �v���p�e�B���܂� LinkFormat �I�u�W�F�N�g��Ԃ��܂��B
    Debug.Print "Decorative: " & shape.Decorative '
'    Debug.Print "FormControlType: " & shape.FormControlType 'XlFormControl
'    Debug.Print "OLEFormat: " & shape.OLEFormat 'OLE �I�u�W�F�N�g �v���p�e�B���܂� OLEFormat �I�u�W�F�N�g
'    Debug.Print "OnAction: " & shape.OnAction '�N���b�N���ꂽ�Ƃ��Ɏ��s����}�N���̖��O��ݒ肵�܂��B

    
    Debug.Print "GraphicStyle: " & shape.GraphicStyle '
'    Debug.Print "GroupItems: " & shape.GroupItems '
    
    Debug.Print "HasChart: " & shape.HasChart '
'    Debug.Print "Chart: " & shape.Chart 'Chart �I�u�W�F�N�g

    Debug.Print "HasSmartArt: " & shape.HasSmartArt '
'    Debug.Print "SmartArt: " & shape.SmartArt '

    Debug.Print "HasTable: " & shape.HasTable '
    If shape.HasTable Then
        For c = 1 To shape.Table.Columns.Count
            For r = 1 To shape.Table.Rows.Count
                If shape.Table.Cell(r, c).Selected Then
                    Debug.Print r & ", " & c
                End If
            Next
        Next
    End If
    
'    Debug.Print "Hyperlink: " & shape.Hyperlink '
    Debug.Print "ID: " & shape.Id '
    Debug.Print "Name: " & shape.Name '
    Debug.Print "Type: " & shape.Type '
    Debug.Print "Title: " & shape.Title '
    Debug.Print "Top: " & shape.Top '
    Debug.Print "Left: " & shape.Left '
    Debug.Print "Width: " & shape.Width '
    Debug.Print "Height: " & shape.Height '
    Debug.Print "HorizontalFlip: " & shape.HorizontalFlip '
    Debug.Print "VerticalFlip: " & shape.VerticalFlip '
    Debug.Print "Rotation: " & shape.Rotation '
    Debug.Print "Visible: " & shape.Visible '
    Debug.Print "ZOrderPosition: " & shape.ZOrderPosition '

    Debug.Print "ConnectionSiteCount: " & shape.ConnectionSiteCount '�w�肳�ꂽ�}�`�̌����_�̐����擾���܂��B
    Debug.Print "Connector: " & shape.Connector 'True �̏ꍇ�A�w�肳�ꂽ�}�`�̓R�l�N�^
'    Debug.Print "ConnectorFormat: " & shape.ConnectorFormat 'ConnectorFormat �I�u�W�F�N�g

'    Debug.Print "Fill: " & shape.Fill '�h��Ԃ��̏����v���p�e�B���i�[���ꂽ FillFormat �I�u�W�F�N�g�܂��� ChartFillFormat �I�u�W�F�N�g
'    Debug.Print "Line: " & shape.Line 'LineFormat �I�u�W�F�N�g

'    Debug.Print "Glow: " & shape.Glow '���ʏ����v���p�e�B���܂ގw�肳�ꂽ�}�`�� GlowFormat �I�u�W�F�N�g
'    Debug.Print "Reflection: " & shape.Reflection '���ˏ����v���p�e�B���܂ގw�肳�ꂽ�}�`�� ReflectionFormat �I�u�W�F�N�g��Ԃ��܂��B
'    Debug.Print "Shadow: " & shape.Shadow '�w�肳�ꂽ�}�`�̉e�̏�����\�� ShadowFormat �I�u�W�F�N�g
'    Debug.Print "Model3D: " & shape.Model3D '
'    Debug.Print "ThreeD: " & shape.ThreeD '3-D ���ʏ����v���p�e�B���܂� ThreeDFormat �I�u�W�F�N�g���擾���܂��B

    
    Debug.Print "LockAspectRatio: " & shape.LockAspectRatio 'True �̏ꍇ�A�w�肳�ꂽ�}�`�́A�T�C�Y��ύX���Ă����̔䗦��ێ����܂��B
    Debug.Print "Locked: " & shape.Locked '

'    Debug.Print "Nodes: " & shape.Nodes '�w�肵���}�`�̊􉽊w�I�ȓ�����\�� ShapeNodes �R���N�V�������擾
'    Debug.Print "PictureFormat: " & shape.PictureFormat '���̃v���p�e�B�́A�}�܂��� OLE �I�u�W�F�N�g��\�� Shape �I�u�W�F�N�g�ɑ΂��Ďg�p���܂��B
    Debug.Print "ShapeStyle: " & shape.ShapeStyle '�}�`�̈�ɂ�����}�`�X�^�C����\�� MsoShapeStyleIndex ���擾�܂��͐ݒ肵�܂��B
'    Debug.Print "SoftEdge: " & shape.SoftEdge '

'    Debug.Print "Vertices: " & shape.Vertices '�t���[�t�H�[���̒��_ (����уx�W�F�Ȑ��̃R���g���[�� �|�C���g) �̍��W����A�̍��W�l?(���W�l: �_�� x ���W�� y ���W��\���l�̃y�A�B���W�̒l�͑����̓_�̒l���܂� 2 �����̔z��Ɋi�[����܂��B)�Ƃ��ĕԂ��܂��B���̃v���p�e�B�ŕԂ��ꂽ�z����AAddCurve ���\�b�h�܂��� AddPolyLine ���\�b�h�̈����Ƃ��Ďw�肷�邱�Ƃ��ł��܂��B

'    Debug.Print "TextEffect: " & shape.TextEffect '�}�`�̓�����ʂ̏����v���p�e�B���܂� TextEffectFormat �I�u�W�F�N�g���擾���܂��B
    If shape.HasTextFrame Then
        If shape.TextFrame.HasText Then
            Debug.Print "TextFrame.TextRange.Text: " & shape.TextFrame.TextRange.text '�}�`�̔z�u����уA���J�[�̃v���p�e�B�̒l���܂� TextFrame �I�u�W�F�N�g��Ԃ��܂��B
            Debug.Print "TextFrame.MarginTop: " & shape.TextFrame.MarginTop '�}�`�̔z�u����уA���J�[�̃v���p�e�B�̒l���܂� TextFrame �I�u�W�F�N�g��Ԃ��܂��B
            Debug.Print "TextFrame.MarginLeft: " & shape.TextFrame.MarginLeft '�}�`�̔z�u����уA���J�[�̃v���p�e�B�̒l���܂� TextFrame �I�u�W�F�N�g��Ԃ��܂��B
            Debug.Print "TextFrame.MarginRight: " & shape.TextFrame.MarginRight '�}�`�̔z�u����уA���J�[�̃v���p�e�B�̒l���܂� TextFrame �I�u�W�F�N�g��Ԃ��܂��B
            Debug.Print "TextFrame.MarginBottom: " & shape.TextFrame.MarginBottom '�}�`�̔z�u����уA���J�[�̃v���p�e�B�̒l���܂� TextFrame �I�u�W�F�N�g��Ԃ��܂��B
'        Debug.Print "TextFrame: " & shape.TextFrame '�}�`�̔z�u����уA���J�[�̃v���p�e�B�̒l���܂� TextFrame �I�u�W�F�N�g��Ԃ��܂��B
'        Debug.Print "TextFrame2: " & shape.TextFrame2 '�}�`�̃e�L�X�g�������܂܂�� TextFrame2 �I�u�W�F�N�g���擾
        End If
    End If

End Sub

Private Sub store(ByRef dictofdict, k, v)
    If Not dictofdict.exists(k) Then
        dictofdict.Add k, v
    Else
        dictofdict(k) = v
    End If
    Debug.Print "store: [" & k & "] = " & v
End Sub

Private Sub storeobj(ByRef dictofdict, k, ByRef v)
    If Not dictofdict.exists(k) Then
        dictofdict.Add k, v
    Else
        Set dictofdict(k) = v
    End If
End Sub


Sub �e�[�u�����̑I���Z���ɃA�N�Z�X�����()
    Debug.Print "HasTable: " & shape.HasTable '
    If shape.HasTable Then
        For c = 1 To shape.Table.Columns.Count
            For r = 1 To shape.Table.Rows.Count
                If shape.Table.Cell(r, c).Selected Then
                    analize_shape shape.Table.Cell(r, c).shape, dictofdict
                End If
            Next
        Next
    End If
End Sub

Private Sub analize_shape_set(ByRef shape, ByRef dictofdict, Optional parentname = "", Optional trow = -1, Optional tcol = -1)
    '���O�����Ă��Ȃ��ꍇ�͓P��
    On Error Resume Next
    prop = shape.Name '�������ɖ��O
    If Err.Number <> 0 Then prop = "#N/A"
    Err.Clear
    On Error GoTo 0
    If prop = "#N/A" Then
        If parentname <> "" And trow <> -1 And tcol <> -1 Then
            shapename = parentname & "[" & trow & "," & tcol & "]"
        Else
            Exit Sub
        End If
    Else
        shapename = prop
    End If
    
    If shape.HasTable Then
        For c = 1 To shape.Table.Columns.Count
            For r = 1 To shape.Table.Rows.Count
                'If shape.Table.Cell(r, c).Selected Then
                    analize_shape_set shape.Table.Cell(r, c).shape, dictofdict, shapename, r, c
                'End If
            Next
        Next
        shepename = prop
    End If
    
    '�I�u�W�F�N�g����
    If Not dictofdict.exists(shapename) Then
        Exit Sub
    End If
    
    Set tgt = dictofdict(shapename)
    
    For Each k In tgt
        v = tgt(k)
        If v = "�" Or v = "�Ǎ��s�\" Or v = "�����s�\" Then
            '
        Else
            On Error Resume Next
            Select Case k
            Case "Top": shape.Top = v
            Case "Left": shape.Left = v
            Case "Width": shape.Width = v
            Case "Height": shape.Height = v
            Case "HorizontalFlip": shape.HorizontalFlip = v
            Case "VerticalFlip": shape.VerticalFlip = v
            Case "Rotation": shape.Rotation = v
            Case "ZOrderPosition": shape.ZOrderPosition = v
            Case "Fill.ForeColor.RGB": shape.Fill.ForeColor.RGB = Val("&H" & v)
            Case "Fill.BackColor.RGB": shape.Fill.BackColor.RGB = Val("&H" & v)
            Case "Fill.Transparency": shape.Fill.Transparency = v
            Case "Line.ForeColor.RGB": shape.Line.ForeColor.RGB = Val("&H" & v)
            Case "Line.BackColor.RGB": shape.Line.BackColor.RGB = Val("&H" & v)
            Case "Line.DashStyle": shape.Line.DashStyle = v
            Case "Line.Weight": shape.Line.Weight = v
            Case "Line.Transparency": shape.Line.Transparency = v
            Case "TextFrame.MarginTop": shape.TextFrame.MarginTop = v
            Case "TextFrame.MarginLeft": shape.TextFrame.MarginLeft = v
            Case "TextFrame.MarginRight": shape.TextFrame.MarginRight = v
            Case "TextFrame.MarginBottom": shape.TextFrame.MarginBottom = v
            End Select
            On Error GoTo 0
        End If
    Next
End Sub
Private Sub analize_shape_get(ByRef shape, ByRef dictofdict, Optional parentname = "", Optional trow = -1, Optional tcol = -1)
    '���O�����Ă��Ȃ��ꍇ�͓P��
    On Error Resume Next
    prop = shape.Name '�������ɖ��O
    If Err.Number <> 0 Then prop = "#N/A"
    Err.Clear
    On Error GoTo 0
    If prop = "#N/A" Then
        If parentname <> "" And trow <> -1 And tcol <> -1 Then
            shapename = parentname & "[" & trow & "," & tcol & "]"
        Else
            Exit Sub
        End If
    Else
        shapename = prop
    End If
    
    If shape.HasTable Then
        For c = 1 To shape.Table.Columns.Count
            For r = 1 To shape.Table.Rows.Count
                'If shape.Table.Cell(r, c).Selected Then
                    analize_shape_get shape.Table.Cell(r, c).shape, dictofdict, shapename, r, c
                'End If
            Next
        Next
        shepename = prop
    End If
    
    '�I�u�W�F�N�g����
    If Not dictofdict.exists(shapename) Then
        dictofdict.Add shapename, CreateObject("scripting.dictionary")
    End If
    
    '���O�͂܂��ŏ��ɐݒ�
    Debug.Print "---------------------" & shapename
    Set tgt = dictofdict(shapename)
    store tgt, "Name", shapename
    If trow <> -1 Then store tgt, "Row", trow
    If tcol <> -1 Then store tgt, "Col", tcol
    store tgt, "TypeName", TypeName(shape)
    
    'ok Debug.Print "BackgroundStyle: " & shape.BackgroundStyle '
    'ok Debug.Print "BlackWhiteMode: " & shape.BlackWhiteMode '
    'ok Debug.Print "GraphicStyle: " & shape.GraphicStyle '

    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.Top '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.Top = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "Top", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.Left '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.Left = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "Left", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.Width '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.Width = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "Width", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.Height '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.Height = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "Height", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.HorizontalFlip '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.HorizontalFlip = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "HorizontalFlip", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.VerticalFlip '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.VerticalFlip = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "VerticalFlip", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.Rotation '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.Rotation = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "Rotation", prop '�����������ɖ��O
    On Error GoTo 0
        
    '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
    On Error Resume Next
    prop = shape.ZOrderPosition '�����������ɖ��O
    If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
    shape.ZOrderPosition = prop '�����������ɖ��O
    If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
    Err.Clear
    store tgt, "ZOrderPosition", prop '�����������ɖ��O
    On Error GoTo 0
    
    If Not shape.HasTable Then
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���iHEX�j
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Fill.ForeColor.RGB), 8) '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Fill.ForeColor.RGB = Val("&H" & prop) '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Fill.ForeColor.RGB", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���iHEX�j
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Fill.BackColor.RGB), 8) '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Fill.BackColor.RGB = Val("&H" & prop) '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Fill.BackColor.RGB", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
        On Error Resume Next
        prop = shape.Fill.Transparency '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Fill.Transparency = prop '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Fill.Transparency", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���iHEX�j
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Line.BackColor.RGB), 8) '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Line.BackColor.RGB = Val("&H" & prop) '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Line.BackColor.RGB", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���iHEX�j
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Line.ForeColor.RGB), 8) '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Line.ForeColor.RGB = Val("&H" & prop) '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Line.ForeColor.RGB", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
        On Error Resume Next
        prop = shape.Line.DashStyle '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Line.DashStyle = prop '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Line.DashStyle", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
        On Error Resume Next
        prop = shape.Line.Weight '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Line.Weight = prop '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Line.Weight", prop '�����������ɖ��O
        On Error GoTo 0
        
        '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
        On Error Resume Next
        prop = shape.Line.Transparency '�����������ɖ��O
        If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
        shape.Line.Transparency = prop '�����������ɖ��O
        If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
        Err.Clear
        store tgt, "Line.Transparency", prop '�����������ɖ��O
        On Error GoTo 0
        
        '    Debug.Print "Glow: " & shape.Glow '���ʏ����v���p�e�B���܂ގw�肳�ꂽ�}�`�� GlowFormat �I�u�W�F�N�g
        '    Debug.Print "Reflection: " & shape.Reflection '���ˏ����v���p�e�B���܂ގw�肳�ꂽ�}�`�� ReflectionFormat �I�u�W�F�N�g��Ԃ��܂��B
        '    Debug.Print "Shadow: " & shape.Shadow '�w�肳�ꂽ�}�`�̉e�̏�����\�� ShadowFormat �I�u�W�F�N�g
        '    Debug.Print "Model3D: " & shape.Model3D '
        '    Debug.Print "ThreeD: " & shape.ThreeD '3-D ���ʏ����v���p�e�B���܂� ThreeDFormat �I�u�W�F�N�g���擾���܂��B
        
        'ok    Debug.Print "ShapeStyle: " & shape.ShapeStyle '�}�`�̈�ɂ�����}�`�X�^�C����\�� MsoShapeStyleIndex ���擾�܂��͐ݒ肵�܂��B
        
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                '�����\�������邩������Ȃ��̂ŕҏW���󂯓���Ȃ�
                On Error Resume Next
                prop = shape.TextFrame.TextRange.text '�����������ɖ��O
                If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
                prop = Replace(Replace(Replace(Replace(Replace(Trim(prop), vbTab, ""), vbLf, ""), vbCr, ""), " ", ""), "�@", "")
                If Len(prop) > 15 Then prop = Left(prop, 13) & "�c�c"
                Err.Clear
                store tgt, "TextFrame.TextRange.Text", prop '�����������ɖ��O
                On Error GoTo 0
        
                '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
                On Error Resume Next
                prop = shape.TextFrame.MarginTop '�����������ɖ��O
                If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
                shape.TextFrame.MarginTop = prop '�����������ɖ��O
                If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
                Err.Clear
                store tgt, "TextFrame.MarginTop", prop '�����������ɖ��O
                On Error GoTo 0
                    
                '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
                On Error Resume Next
                prop = shape.TextFrame.MarginLeft '�����������ɖ��O
                If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
                shape.TextFrame.MarginLeft = prop '�����������ɖ��O
                If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
                Err.Clear
                store tgt, "TextFrame.MarginLeft", prop '�����������ɖ��O
                On Error GoTo 0
                    
                '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
                On Error Resume Next
                prop = shape.TextFrame.MarginRight '�����������ɖ��O
                If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
                shape.TextFrame.MarginRight = prop '�����������ɖ��O
                If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
                Err.Clear
                store tgt, "TextFrame.MarginRight", prop '�����������ɖ��O
                On Error GoTo 0
                    
                '���݂��Ȃ������m��Ȃ��v���p�e�B�������ꍇ�̃p�^�[���i���l�j
                On Error Resume Next
                prop = shape.TextFrame.MarginBottom '�����������ɖ��O
                If Err.Number <> 0 Then prop = "�Ǎ��s�\" '�ǂݍ��ݕs�\�̏ꍇ
                shape.TextFrame.MarginBottom = prop '�����������ɖ��O
                If Err.Number <> 0 And prop <> "�Ǎ��s�\" Then prop = "�����s�\"
                Err.Clear
                store tgt, "TextFrame.MarginBottom", prop '�����������ɖ��O
                On Error GoTo 0
            End If
        End If
    End If

End Sub


Private Sub analize_shape_findtable(ByRef shape, ByRef dictofdict)
    If shape.HasTable Then
        dictofdict.Add shape.Name, shape
    End If
End Sub



Private Sub analize_shape_hastextframe(ByRef shape, ByRef dictofdict)
    If shape.HasTextFrame Then
        dictofdict.Add shape.Name, shape
    End If
End Sub



Private Sub analize_shape_selector(ByRef shape, action, ByRef dictofdict)
    Select Case action
    Case "get"
        analize_shape_get shape, dictofdict
    Case "set"
        analize_shape_set shape, dictofdict
    Case "findtable"
        analize_shape_findtable shape, dictofdict
    Case "hastextframe"
        analize_shape_hastextframe shape, dictofdict
    Case Else
        MsgBox "undefined"
        End
    End Select
End Sub



Private Sub analize_shapeorgroup(ByRef shape, action, ByRef dictofdict)
    
    '�}�`���O���[�v�����Ă��邩����
    If shape.Type = msoGroup Then
        '��2�K�w�̐}�`�����[�v
        For Each b In shape.GroupItems
            analize_shape_selector b, action, dictofdict
        Next
    Else
        analize_shape_selector shape, action, dictofdict
    End If
    
End Sub

Private Sub analize_shaperange(ByRef shaperange, action, ByRef dictofdict)
    
    If shaperange.Type = msoGroup Then
        For Each b In shaperange.GroupItems
            analize_shapeorgroup b, action, dictofdict
        Next
    Else
        For Each b In shaperange
            analize_shapeorgroup b, action, dictofdict
        Next
    End If
    
End Sub

Function analize_selection(action)

    With ActiveWindow.Selection
        If .Type >= ppSelectionShapes Then
            Set dictofdict = CreateObject("scripting.dictionary")
            analize_shaperange .shaperange, action, dictofdict
            Set analize_selection = dictofdict
            Set dictofdict = Nothing
        Else
            Set analize_selection = Nothing
        End If
    End With

End Function




Sub analize_selection_preset(action, dictofdict)

    With ActiveWindow.Selection
        If .Type >= ppSelectionShapes Then
            analize_shaperange .shaperange, action, dictofdict
        End If
    End With

End Sub





Sub �I���z�o()

    With ActiveWindow.Selection
        If .Type = ppSelectionNone Then
            Debug.Print "�����I������Ă��܂���"
        End If
        If .Type >= ppSelectionSlides Then
            Debug.Print "�X���C�h���I������Ă��܂�"
        End If
        If .Type >= ppSelectionShapes Then
            Debug.Print "�V�F�C�v���I������Ă��܂�"
            Set dictofdict = CreateObject("scripting.dictionary")
            analize_shaperange .shaperange, "", dictofdict
            Set dictofdict = Nothing
        End If
        If .Type >= ppSelectionText Then
            Debug.Print "�e�L�X�g�͈͂��I������Ă��܂�"
        End If
    End With

End Sub
