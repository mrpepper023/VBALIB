Attribute VB_Name = "PptMisc"
'64bit��
Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )

'32bit��
'Private Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )

Public Sub �S�̂��B��()
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
End Sub

Public Sub �A�N�e�B�u��ʂ��B��()
    keybd_event &HA4, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    keybd_event &HA4, 0&, &H1 Or &H2, 0&
End Sub




Sub �}�`�̒��_�����炩�ɂ���()

    With ActiveWindow.Selection.shaperange.Nodes

        For i = .Count To 1 Step -1

            .SetEditingType i, msoEditingSmooth
            .SetEditingType i, msoSegmentLine
            
        Next i

    End With
End Sub



' n���]�Ώ�

'

Sub edit_freeform()

    
    With ActiveWindow.Selection.shaperange.Nodes
        pointsArray = .Item(2).Points
        currXvalue = pointsArray(1, 1)
        currYvalue = pointsArray(1, 2)
        .SetPosition 2, currXvalue + 200, currYvalue + 300
    End With

End Sub

Sub create_freeform()
    
    Set myDocument = ActivePresentation.Slides(1)
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=360, Y1:=200)
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, X1:=380, Y1:=230, X2:=400, Y2:=250, X3:=450, Y3:=300
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, X1:=480, Y1:=300, X2:=480, Y2:=300, X3:=480, Y3:=200 '�E��
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingCorner, X1:=480, Y1:=400 '�ŉ�
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingCorner, X1:=360, Y1:=200
        .ConvertToShape
    End With
    
End Sub
