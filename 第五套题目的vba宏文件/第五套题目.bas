Attribute VB_Name = "NewMacros"
Sub ��һ��()
Attribute ��һ��.VB_Description = "��һ��"
Attribute ��һ��.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.��һ��"
'
' ��һ�� ��
' ��һ��
'

Dim score As Integer
score = 0

    Selection.WholeStory
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        
        'ҳ��߾������Ϊ����Ϊ3
'test 0
If ActiveDocument.PageSetup.TopMargin - CentimetersToPoints(3) < 0.1 Then
   score = score + 2
End If

        
                
                .TopMargin = CentimetersToPoints(3)
                .BottomMargin = CentimetersToPoints(3)
                .LeftMargin = CentimetersToPoints(3)
                .RightMargin = CentimetersToPoints(3)
                
                
                .Gutter = CentimetersToPoints(0)
                .HeaderDistance = CentimetersToPoints(1.5)
                .FooterDistance = CentimetersToPoints(1.75)
                
                'ҳ������ΪA4ֽ�ţ���Լ��8.27Ӣ�� �� 11.75Ӣ�磨21���� �� 29.7���ף�
'test 1
'MsgBox ActiveDocument.PageSetup.PageWidth & CentimetersToPoints(21)
If ActiveDocument.PageSetup.PageWidth - CentimetersToPoints(21) < 0.1 Then
  MsgBox ture
  score = score + 2
Else
    MsgBox flase
End If
  


        
                .PageWidth = CentimetersToPoints(21)
                .PageHeight = CentimetersToPoints(29.7)
                
                .FirstPageTray = wdPrinterDefaultBin
                .OtherPagesTray = wdPrinterDefaultBin
                .SectionStart = wdSectionNewPage
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .VerticalAlignment = wdAlignVerticalTop
                .SuppressEndnotes = False
                .MirrorMargins = False
                .TwoPagesOnOne = False
                .BookFoldPrinting = False
                .BookFoldRevPrinting = False
                .BookFoldPrintingSheets = 1
                .GutterPos = wdGutterPosLeft
                .LayoutMode = wdLayoutModeLineGrid
                
        '��һ���֣��������ˮӡ��
    End With
    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    'AddTextEffect�� PresetTextEffect ���ı����������ơ��ֺš� FontBold ����б���塢��ࡢ������
'test 3
'���ˮӡ�޷������ж�
                Selection.HeaderFooter.Shapes.AddTextEffect( _
                    PowerPlusWaterMarkObject79406343, "��������", "Times New Roman", 1, False, _
                    False, 0, 0).Select
                '��֪��������Ϊʲô���в��ˣ�ע��������ɿ�������
                '    Selection.ShapeRange.Name = "PowerPlusWaterMarkObject79406343"
                Selection.ShapeRange.TextEffect.NormalizedHeight = False
                Selection.ShapeRange.Line.Visible = False
                Selection.ShapeRange.Fill.Visible = True
                Selection.ShapeRange.Fill.Solid
                Selection.ShapeRange.Fill.ForeColor.RGB = RGB(192, 192, 192)
                Selection.ShapeRange.Fill.Transparency = 0.5
                Selection.ShapeRange.Rotation = 315
                Selection.ShapeRange.LockAspectRatio = True
                Selection.ShapeRange.Height = CentimetersToPoints(4.23)
                Selection.ShapeRange.Width = CentimetersToPoints(16.92)
                Selection.ShapeRange.WrapFormat.AllowOverlap = True
                Selection.ShapeRange.WrapFormat.Side = wdWrapNone
                Selection.ShapeRange.WrapFormat.Type = 3
                Selection.ShapeRange.RelativeHorizontalPosition = _
                    wdRelativeVerticalPositionMargin
                Selection.ShapeRange.RelativeVerticalPosition = _
                    wdRelativeVerticalPositionMargin
                'Ϊ�ĵ����ҳü
'test 4  ���۵�����ҳ�沼�֣�����ҳ��ı߾࣬���ҳ�沼���Ƿ���Ĺ������Ĺ�����ú����ӷ�
    
    
    
'test 5     ҳü�����ؼ���
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.TypeText Text:="��������"
'test 6
                '������������Ϊ
                 Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.62)

End Sub
