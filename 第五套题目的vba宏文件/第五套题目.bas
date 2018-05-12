Attribute VB_Name = "NewMacros"
Sub 第一题()
Attribute 第一题.VB_Description = "第一题"
Attribute 第一题.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.第一题"
'
' 第一题 宏
' 第一题
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
        
        '页面边距的设置为各个为3
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
                
                '页面设置为A4纸张，大约在8.27英寸 × 11.75英寸（21厘米 × 29.7厘米）
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
                
        '第一部分，设计文字水印，
    End With
    ActiveDocument.Sections(1).Range.Select
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    'AddTextEffect（ PresetTextEffect 、文本、字体名称、字号、 FontBold 、倾斜字体、左侧、顶部）
'test 3
'添加水印无法进行判断
                Selection.HeaderFooter.Shapes.AddTextEffect( _
                    PowerPlusWaterMarkObject79406343, "空气环境", "Times New Roman", 1, False, _
                    False, 0, 0).Select
                '不知道这个语句为什么运行不了，注解掉了依旧可以运行
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
                '为文档添加页眉
'test 4  主观的设置页面布局，设置页面的边距，检测页面布局是否更改过，更改过则调用函数加分
    
    
    
'test 5     页眉包含关键字
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.TypeText Text:="空气环境"
'test 6
                '设置首行缩进为
                 Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.62)

End Sub
