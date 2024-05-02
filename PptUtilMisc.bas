Attribute VB_Name = "PptMisc"
'64bit版
Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )

'32bit版
'Private Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
        )

Public Sub 全体を撮る()
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
End Sub

Public Sub アクティブ画面を撮る()
    keybd_event &HA4, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1, 0&
    keybd_event vbKeySnapshot, 0&, &H1 Or &H2, 0&
    keybd_event &HA4, 0&, &H1 Or &H2, 0&
End Sub




Sub 図形の頂点を滑らかにする()

    With ActiveWindow.Selection.shaperange.Nodes

        For i = .Count To 1 Step -1

            .SetEditingType i, msoEditingSmooth
            .SetEditingType i, msoSegmentLine
            
        Next i

    End With
End Sub



' n回回転対称

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
        .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, X1:=480, Y1:=300, X2:=480, Y2:=300, X3:=480, Y3:=200 '右上
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingCorner, X1:=480, Y1:=400 '最下
        .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingCorner, X1:=360, Y1:=200
        .ConvertToShape
    End With
    
End Sub
