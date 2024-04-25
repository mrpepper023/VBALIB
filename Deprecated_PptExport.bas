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
'    Debug.Print "Placement: " & shape.Placement 'オブジェクトとその下にあるセルとの位置関係を表す XlPlacement

'    Debug.Print "Callout: " & shape.Callout '吹き出しの書式プロパティ
    Debug.Print "Child: " & shape.Child '
'    Debug.Print "Parent: " & shape.Parent '
'    Debug.Print "ParentGroup: " & shape.ParentGroup '
    
'    Debug.Print "ControlFormat: " & shape.ControlFormat 'for excel?
    Debug.Print "Application: " & shape.Application '指定されたオブジェクトを作成した Application オブジェクトを返します。
    Debug.Print "Creator: " & shape.Creator 'オブジェクトが作成されたアプリケーションを示す 32 ビットの整数
'    Debug.Print "LinkFormat: " & shape.LinkFormat 'リンクされた OLE オブジェクト プロパティを含む LinkFormat オブジェクトを返します。
    Debug.Print "Decorative: " & shape.Decorative '
'    Debug.Print "FormControlType: " & shape.FormControlType 'XlFormControl
'    Debug.Print "OLEFormat: " & shape.OLEFormat 'OLE オブジェクト プロパティを含む OLEFormat オブジェクト
'    Debug.Print "OnAction: " & shape.OnAction 'クリックされたときに実行するマクロの名前を設定します。

    
    Debug.Print "GraphicStyle: " & shape.GraphicStyle '
'    Debug.Print "GroupItems: " & shape.GroupItems '
    
    Debug.Print "HasChart: " & shape.HasChart '
'    Debug.Print "Chart: " & shape.Chart 'Chart オブジェクト

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

    Debug.Print "ConnectionSiteCount: " & shape.ConnectionSiteCount '指定された図形の結合点の数を取得します。
    Debug.Print "Connector: " & shape.Connector 'True の場合、指定された図形はコネクタ
'    Debug.Print "ConnectorFormat: " & shape.ConnectorFormat 'ConnectorFormat オブジェクト

'    Debug.Print "Fill: " & shape.Fill '塗りつぶしの書式プロパティが格納された FillFormat オブジェクトまたは ChartFillFormat オブジェクト
'    Debug.Print "Line: " & shape.Line 'LineFormat オブジェクト

'    Debug.Print "Glow: " & shape.Glow '光彩書式プロパティを含む指定された図形の GlowFormat オブジェクト
'    Debug.Print "Reflection: " & shape.Reflection '反射書式プロパティを含む指定された図形の ReflectionFormat オブジェクトを返します。
'    Debug.Print "Shadow: " & shape.Shadow '指定された図形の影の書式を表す ShadowFormat オブジェクト
'    Debug.Print "Model3D: " & shape.Model3D '
'    Debug.Print "ThreeD: " & shape.ThreeD '3-D 効果書式プロパティを含む ThreeDFormat オブジェクトを取得します。

    
    Debug.Print "LockAspectRatio: " & shape.LockAspectRatio 'True の場合、指定された図形は、サイズを変更しても元の比率を保持します。
    Debug.Print "Locked: " & shape.Locked '

'    Debug.Print "Nodes: " & shape.Nodes '指定した図形の幾何学的な特徴を表す ShapeNodes コレクションを取得
'    Debug.Print "PictureFormat: " & shape.PictureFormat 'このプロパティは、図または OLE オブジェクトを表す Shape オブジェクトに対して使用します。
    Debug.Print "ShapeStyle: " & shape.ShapeStyle '図形領域における図形スタイルを表す MsoShapeStyleIndex を取得または設定します。
'    Debug.Print "SoftEdge: " & shape.SoftEdge '

'    Debug.Print "Vertices: " & shape.Vertices 'フリーフォームの頂点 (およびベジェ曲線のコントロール ポイント) の座標を一連の座標値?(座標値: 点の x 座標と y 座標を表す値のペア。座標の値は多くの点の値を含む 2 次元の配列に格納されます。)として返します。このプロパティで返された配列を、AddCurve メソッドまたは AddPolyLine メソッドの引数として指定することができます。

'    Debug.Print "TextEffect: " & shape.TextEffect '図形の特殊効果の書式プロパティを含む TextEffectFormat オブジェクトを取得します。
    If shape.HasTextFrame Then
        If shape.TextFrame.HasText Then
            Debug.Print "TextFrame.TextRange.Text: " & shape.TextFrame.TextRange.text '図形の配置およびアンカーのプロパティの値を含む TextFrame オブジェクトを返します。
            Debug.Print "TextFrame.MarginTop: " & shape.TextFrame.MarginTop '図形の配置およびアンカーのプロパティの値を含む TextFrame オブジェクトを返します。
            Debug.Print "TextFrame.MarginLeft: " & shape.TextFrame.MarginLeft '図形の配置およびアンカーのプロパティの値を含む TextFrame オブジェクトを返します。
            Debug.Print "TextFrame.MarginRight: " & shape.TextFrame.MarginRight '図形の配置およびアンカーのプロパティの値を含む TextFrame オブジェクトを返します。
            Debug.Print "TextFrame.MarginBottom: " & shape.TextFrame.MarginBottom '図形の配置およびアンカーのプロパティの値を含む TextFrame オブジェクトを返します。
'        Debug.Print "TextFrame: " & shape.TextFrame '図形の配置およびアンカーのプロパティの値を含む TextFrame オブジェクトを返します。
'        Debug.Print "TextFrame2: " & shape.TextFrame2 '図形のテキスト書式が含まれる TextFrame2 オブジェクトを取得
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


Sub テーブル内の選択セルにアクセスする例()
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
    '名前がついていない場合は撤退
    On Error Resume Next
    prop = shape.Name '※ここに名前
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
    
    'オブジェクト特定
    If Not dictofdict.exists(shapename) Then
        Exit Sub
    End If
    
    Set tgt = dictofdict(shapename)
    
    For Each k In tgt
        v = tgt(k)
        If v = "蟲" Or v = "読込不能" Or v = "書込不能" Then
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
    '名前がついていない場合は撤退
    On Error Resume Next
    prop = shape.Name '※ここに名前
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
    
    'オブジェクト生成
    If Not dictofdict.exists(shapename) Then
        dictofdict.Add shapename, CreateObject("scripting.dictionary")
    End If
    
    '名前はまず最初に設定
    Debug.Print "---------------------" & shapename
    Set tgt = dictofdict(shapename)
    store tgt, "Name", shapename
    If trow <> -1 Then store tgt, "Row", trow
    If tcol <> -1 Then store tgt, "Col", tcol
    store tgt, "TypeName", TypeName(shape)
    
    'ok Debug.Print "BackgroundStyle: " & shape.BackgroundStyle '
    'ok Debug.Print "BlackWhiteMode: " & shape.BlackWhiteMode '
    'ok Debug.Print "GraphicStyle: " & shape.GraphicStyle '

    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.Top '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.Top = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "Top", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.Left '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.Left = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "Left", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.Width '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.Width = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "Width", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.Height '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.Height = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "Height", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.HorizontalFlip '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.HorizontalFlip = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "HorizontalFlip", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.VerticalFlip '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.VerticalFlip = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "VerticalFlip", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.Rotation '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.Rotation = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "Rotation", prop '◆◆◆ここに名前
    On Error GoTo 0
        
    '存在しないかも知れないプロパティを試す場合のパターン（直値）
    On Error Resume Next
    prop = shape.ZOrderPosition '◆◆◆ここに名前
    If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
    shape.ZOrderPosition = prop '◆◆◆ここに名前
    If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
    Err.Clear
    store tgt, "ZOrderPosition", prop '◆◆◆ここに名前
    On Error GoTo 0
    
    If Not shape.HasTable Then
        '存在しないかも知れないプロパティを試す場合のパターン（HEX）
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Fill.ForeColor.RGB), 8) '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Fill.ForeColor.RGB = Val("&H" & prop) '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Fill.ForeColor.RGB", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（HEX）
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Fill.BackColor.RGB), 8) '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Fill.BackColor.RGB = Val("&H" & prop) '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Fill.BackColor.RGB", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（直値）
        On Error Resume Next
        prop = shape.Fill.Transparency '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Fill.Transparency = prop '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Fill.Transparency", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（HEX）
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Line.BackColor.RGB), 8) '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Line.BackColor.RGB = Val("&H" & prop) '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Line.BackColor.RGB", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（HEX）
        On Error Resume Next
        prop = Right("00000000" & Hex(shape.Line.ForeColor.RGB), 8) '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Line.ForeColor.RGB = Val("&H" & prop) '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Line.ForeColor.RGB", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（直値）
        On Error Resume Next
        prop = shape.Line.DashStyle '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Line.DashStyle = prop '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Line.DashStyle", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（直値）
        On Error Resume Next
        prop = shape.Line.Weight '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Line.Weight = prop '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Line.Weight", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '存在しないかも知れないプロパティを試す場合のパターン（直値）
        On Error Resume Next
        prop = shape.Line.Transparency '◆◆◆ここに名前
        If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
        shape.Line.Transparency = prop '◆◆◆ここに名前
        If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
        Err.Clear
        store tgt, "Line.Transparency", prop '◆◆◆ここに名前
        On Error GoTo 0
        
        '    Debug.Print "Glow: " & shape.Glow '光彩書式プロパティを含む指定された図形の GlowFormat オブジェクト
        '    Debug.Print "Reflection: " & shape.Reflection '反射書式プロパティを含む指定された図形の ReflectionFormat オブジェクトを返します。
        '    Debug.Print "Shadow: " & shape.Shadow '指定された図形の影の書式を表す ShadowFormat オブジェクト
        '    Debug.Print "Model3D: " & shape.Model3D '
        '    Debug.Print "ThreeD: " & shape.ThreeD '3-D 効果書式プロパティを含む ThreeDFormat オブジェクトを取得します。
        
        'ok    Debug.Print "ShapeStyle: " & shape.ShapeStyle '図形領域における図形スタイルを表す MsoShapeStyleIndex を取得または設定します。
        
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                '内部構造があるかもしれないので編集を受け入れない
                On Error Resume Next
                prop = shape.TextFrame.TextRange.text '◆◆◆ここに名前
                If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
                prop = Replace(Replace(Replace(Replace(Replace(Trim(prop), vbTab, ""), vbLf, ""), vbCr, ""), " ", ""), "　", "")
                If Len(prop) > 15 Then prop = Left(prop, 13) & "……"
                Err.Clear
                store tgt, "TextFrame.TextRange.Text", prop '◆◆◆ここに名前
                On Error GoTo 0
        
                '存在しないかも知れないプロパティを試す場合のパターン（直値）
                On Error Resume Next
                prop = shape.TextFrame.MarginTop '◆◆◆ここに名前
                If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
                shape.TextFrame.MarginTop = prop '◆◆◆ここに名前
                If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
                Err.Clear
                store tgt, "TextFrame.MarginTop", prop '◆◆◆ここに名前
                On Error GoTo 0
                    
                '存在しないかも知れないプロパティを試す場合のパターン（直値）
                On Error Resume Next
                prop = shape.TextFrame.MarginLeft '◆◆◆ここに名前
                If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
                shape.TextFrame.MarginLeft = prop '◆◆◆ここに名前
                If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
                Err.Clear
                store tgt, "TextFrame.MarginLeft", prop '◆◆◆ここに名前
                On Error GoTo 0
                    
                '存在しないかも知れないプロパティを試す場合のパターン（直値）
                On Error Resume Next
                prop = shape.TextFrame.MarginRight '◆◆◆ここに名前
                If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
                shape.TextFrame.MarginRight = prop '◆◆◆ここに名前
                If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
                Err.Clear
                store tgt, "TextFrame.MarginRight", prop '◆◆◆ここに名前
                On Error GoTo 0
                    
                '存在しないかも知れないプロパティを試す場合のパターン（直値）
                On Error Resume Next
                prop = shape.TextFrame.MarginBottom '◆◆◆ここに名前
                If Err.Number <> 0 Then prop = "読込不能" '読み込み不能の場合
                shape.TextFrame.MarginBottom = prop '◆◆◆ここに名前
                If Err.Number <> 0 And prop <> "読込不能" Then prop = "書込不能"
                Err.Clear
                store tgt, "TextFrame.MarginBottom", prop '◆◆◆ここに名前
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
    
    '図形がグループ化しているか判定
    If shape.Type = msoGroup Then
        '第2階層の図形をループ
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





Sub 選択吸出()

    With ActiveWindow.Selection
        If .Type = ppSelectionNone Then
            Debug.Print "何も選択されていません"
        End If
        If .Type >= ppSelectionSlides Then
            Debug.Print "スライドが選択されています"
        End If
        If .Type >= ppSelectionShapes Then
            Debug.Print "シェイプが選択されています"
            Set dictofdict = CreateObject("scripting.dictionary")
            analize_shaperange .shaperange, "", dictofdict
            Set dictofdict = Nothing
        End If
        If .Type >= ppSelectionText Then
            Debug.Print "テキスト範囲が選択されています"
        End If
    End With

End Sub
