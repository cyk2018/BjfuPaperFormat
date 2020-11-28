Attribute VB_Name = "realword"
Sub 插入分节符()
    Selection.InsertBreak Type:=wdSectionBreakNextPage
End Sub
Sub 插入分页符()
 Selection.InsertBreak Type:=wdPageBreak
End Sub
Sub 中文全文表格设置()
    response = MsgBox("【重要提示】" & Chr(13) & "确认对全文所有表格进行设置吗？" & Chr(13) & "（包括 1 x 1 表格）", buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
        Exit Sub
    End If
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        '选中表格
        tb.Select
        '设置表格的格式
        'If tb.Title = "" Then
        
        '设置内外边框
        tb.Borders.InsideLineStyle = wdLineStyleSingle
        tb.Borders.InsideLineWidth = wdLineWidth050pt
        tb.Borders.OutsideLineStyle = wdLineStyleSingle
        tb.Borders.OutsideLineWidth = wdLineWidth050pt
        
        tb.Style.Font.Name = "宋体"
        tb.Style.Font.Size = 10.5
        tb.Style.Font.NameAscii = "Times New Roman"
        
        
        '标题设置
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "表 " & "x-x 请输入标题"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '表题采用宋体五号
        Selection.Font.Name = "宋体"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
        Selection.Font.Bold = False
                       
        'End If
        
    Next i
    MsgBox "请手动输入标题"
End Sub
Sub 中文选中表格设置()
    '选中表格说明此属性为5
    If Selection.Type <> 5 Then
        MsgBox "当前尚未选中表格"
        Exit Sub
    End If
    '设置内外边框
        Selection.Borders.InsideLineStyle = wdLineStyleSingle
        Selection.Borders.InsideLineWidth = wdLineWidth050pt
        Selection.Borders.OutsideLineStyle = wdLineStyleSingle
        Selection.Borders.OutsideLineWidth = wdLineWidth050pt
        
        Selection.Style.Font.Name = "宋体"
        Selection.Style.Font.Size = 10.5
        Selection.Style.Font.NameAscii = "Times New Roman"
        
        
        '标题设置
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "表 " & "x-x 请输入标题"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '表题采用宋体五号
        Selection.Font.Name = "宋体"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
        Selection.Font.Bold = False
    MsgBox "请手动输入标题"
End Sub
Sub 英文全文表格设置()
    response = MsgBox("【重要提示】" & Chr(13) & "确认对全文所有表格进行设置吗？" & Chr(13) & "（包括 1 x 1 表格）", buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
        Exit Sub
    End If
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        '选中表格
        tb.Select
        '设置表格的格式
        'If tb.Title = "" Then
        
        '设置内外边框
        tb.Borders.InsideLineStyle = wdLineStyleSingle
        tb.Borders.InsideLineWidth = wdLineWidth050pt
        tb.Borders.OutsideLineStyle = wdLineStyleSingle
        tb.Borders.OutsideLineWidth = wdLineWidth050pt
        
        tb.Style.Font.Name = "宋体"
        tb.Style.Font.Size = 10.5
        tb.Style.Font.NameAscii = "Times New Roman"
        
        
        '标题设置
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "Table" & " x-x Please input the title"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '表题采用宋体五号
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Bold = False
        
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
                    
        'End If
        
    Next i
    MsgBox "Please enter the title manually"
End Sub
Sub 英文选中表格设置()
    '选中表格说明此属性为5
    If Selection.Type <> 5 Then
        MsgBox "当前尚未选中表格"
        Exit Sub
    End If
    '设置内外边框
        Selection.Borders.InsideLineStyle = wdLineStyleSingle
        Selection.Borders.InsideLineWidth = wdLineWidth050pt
        Selection.Borders.OutsideLineStyle = wdLineStyleSingle
        Selection.Borders.OutsideLineWidth = wdLineWidth050pt
        
        Selection.Style.Font.Name = "宋体"
        Selection.Style.Font.Size = 10.5
        Selection.Style.Font.NameAscii = "Times New Roman"
        
        
        '标题设置
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "Table" & " x-x Please input the title"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Bold = False
        
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
    MsgBox "请手动输入标题"
End Sub

Sub 中文全文图片设置()
    response = MsgBox("【重要提示】" & Chr(13) & "确认对全文所有表格进行设置吗？" & Chr(13) & "（可能包括公式中的图片）", buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
        Exit Sub
    End If
    Dim pics As InlineShapes, pic As InlineShape
    Set pics = ActiveDocument.InlineShapes
    n = pics.Count
    For i = 1 To n
        Set pic = pics(i)
        pic.Select
        'If pic.Title = "" Then
            Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveRight
            pictitle = Chr(13) & "图 x-x" & " 请输入标题"
            Selection.TypeText pictitle
            'pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "宋体"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
       ' End If
    Next i
    MsgBox "请手动输入标题"
End Sub
Sub 中文选中图片设置()
    '选中表格说明此属性为5
    If Selection.Type = wdSelectionShape Or Selection.Type = wdSelectionShape Then
        MsgBox "当前尚未选中图片"
        Exit Sub
    End If
    Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveRight
            pictitle = Chr(13) & "图 x-x" & " 请输入标题"
            Selection.TypeText pictitle
            'pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "宋体"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
End Sub


Sub 英文全文图片设置()
    response = MsgBox("【重要提示】" & Chr(13) & "确认对全文所有表格进行设置吗？" & Chr(13) & "（可能包括公式中的图片）", buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
        Exit Sub
    End If
    Dim pics As InlineShapes, pic As InlineShape
    Set pics = ActiveDocument.InlineShapes
    n = pics.Count
    For i = 1 To n
        Set pic = pics(i)
        pic.Select
        'If pic.Title = "" Then
            Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveDown unit:=wdParagraph, Count:=1, Extend:=wdExtend
            pictitle = Chr(13) & "Figure x-x" & " Please input the title"
            Selection.TypeText pictitle
            'pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "Times New Roman"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
        'End If
    Next i
    MsgBox "Please enter the title manually"
End Sub

Sub 英文选中图片设置()
    '选中表格说明此属性为5
    If Selection.Type = wdSelectionShape Or Selection.Type = wdSelectionShape Then
        MsgBox "当前尚未选中图片"
        Exit Sub
    End If
    Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveDown unit:=wdParagraph, Count:=1, Extend:=wdExtend
            pictitle = Chr(13) & "Figure x-x" & " Please input the title"
            Selection.TypeText pictitle
            'pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "Times New Roman"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
End Sub



Sub 全文页边距和页眉_页脚格式()
response = MsgBox("【重要提示】" & Chr(13) & "确认后会覆盖已修改的页眉" & Chr(13) & "请谨慎点击确定", buttons:=vbOKCancel + vbDefaultButton2)
    If response = 1 Then
        response = MsgBox("【再次确认】" & Chr(13) & "你确定修改吗？", buttons:=vbOKCancel + vbDefaultButton2)
        If response <> 1 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If


    Selection.WholeStory
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(3)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(3)
    End With
   
    
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
  
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Delete
    Selection.TypeText Text:="电子科技大学学士学位论文 "
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 10.5
   

    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Delete
    Selection.TypeText Text:=""
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 10.5
    
    Selection.PageSetup.HeaderDistance = 56.6
    Selection.PageSetup.FooterDistance = 56.6 ' 设置页脚到页面底边的距离
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.HomeKey
     Dim oWord As Word.Application
    Set oWord = Word.Application
    Dim oDoc As Document
    Dim oSec As Section
    Dim oFoot As HeaderFooter
    Dim oHead As HeaderFooter
    Set oDoc = oWord.ActiveDocument
     With oDoc
        'iCount = .BuiltInDocumentProperties(wdPropertyPages)
        iCount = .Sections.Count
        For i = 1 To iCount
            Set oSec = .Sections(i)
            With oSec
             
                '页眉
                Set oHead = .Headers(wdHeaderFooterPrimary)
                '页脚
                Set oFoot = .Footers(wdHeaderFooterPrimary)
                '页眉取消链接到前一节
                oHead.LinkToPrevious = False
                '页脚到前一节
                oFoot.LinkToPrevious = False
            
            End With
        Next i
    End With
    MsgBox "完成！"
End Sub

Sub 英文数字_页眉格式()
     If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定文字！"
        Exit Sub
    End If
    response = MsgBox("此操作用于设置英文数字页眉的格式" & Chr(13), buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
            Exit Sub
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '设置页眉
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 10.5
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub


Sub 全文英文使用新罗马字体()
    On Error GoTo msg
    Dim ps As Paragraphs
    Set ps = ActiveDocument.Paragraphs
    n = ps.Count
    For i = 1 To n
        For Each C In ps(i).Range.Characters
            If VBA.Asc(C) >= 0 And C.Font.Name <> "Times New Roman" Then
                C.Font.Name = "Times New Roman"
            End If
        Next
    Next
    MsgBox "完成！"
    Exit Sub
msg:
    MsgBox "出了点问题，请检查后重试  >_<", Title:="Error", buttons:=vbCritical
End Sub

Sub 正文_中文()
    If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定文字！"
        Exit Sub
    End If
    Selection.ClearFormatting
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 12
    Selection.Font.Bold = False
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 20 '固定行距20磅
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        '.LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 2 '首行缩进两字符
        .OutlineLevel = wdOutlineLevelBodyText
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub
Sub 正文_英文()
    If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定英文！"
        Exit Sub
    End If
    Selection.ClearFormatting
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    
    Selection.Font.Bold = False
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 20
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        '.LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 2 '首行缩进两字符
        .OutlineLevel = wdOutlineLevelBodyText
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub

Sub 生成页码()
    response = MsgBox("【重要提示】" & Chr(13) & "请确保光标位于第一章的一级标题前" & Chr(13) & "再点击[确定]  否则点击[取消]", buttons:=vbOKCancel + vbDefaultButton2)
    If response = 1 Then
        response = MsgBox("【再次确认】" & Chr(13) & "我确认光标位于正确位置", buttons:=vbOKCancel + vbDefaultButton2)
        If response <> 1 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    Selection.InsertBreak Type:=wdSectionBreakNextPage
    
    With ActiveDocument.Sections(1)
        .Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
        .Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        .Footers(wdHeaderFooterPrimary).PageNumbers.NumberStyle = wdPageNumberStyleUppercaseRoman '第一章之前用新罗马字体
        .Footers(wdHeaderFooterPrimary).Range.Font.Name = "Times New Roman"
        .Footers(wdHeaderFooterPrimary).Range.Font.Size = 9
        '.Headers(1).Range.ParagraphFormat.Borders(3).LineStyle = wdLineStyleNone
    End With
    With ActiveDocument.Sections(2)
        .Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=True
        .Footers(wdHeaderFooterPrimary).PageNumbers.RestartNumberingAtSection = True
        .Footers(wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        .Footers(wdHeaderFooterPrimary).PageNumbers.NumberStyle = wdPageNumberStyleArabic
        .Footers(wdHeaderFooterPrimary).Range.Font.Name = "Times New Roman"
        .Footers(wdHeaderFooterPrimary).Range.Font.Size = 9
        '.Headers(1).Range.ParagraphFormat.Borders(3).LineStyle = wdLineStyleNone
    End With
    MsgBox "完成！"
End Sub


Sub 插入脚注()
    With Selection
        With .FootnoteOptions
            .Location = wdBottomOfPage
            '.NumberingRule = wdRestartSection
            '.NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleNumberInCircle
            .LayoutColumns = 1
        End With
        .Font.Size = 10.5
        .ParagraphFormat.Alignment = 3
        .ParagraphFormat.LineSpacingRule = pbLineSpacingSingle '单倍行距
        .ParagraphFormat.LineUnitBefore = 0 '段前距
        .ParagraphFormat.LineUnitAfter = 0 '段后距
        Set myRange = ActiveDocument.Sections(1).Range
    If myRange.Footnotes.NumberingRule = wdRestartSection Then
        myRange.Footnotes.NumberingRule = wdRestartPage
    End If
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
End Sub

Sub 中文章标题()
    With ActiveDocument.Styles("标题 1").Font
        .NameFarEast = "黑体"
        .Name = "黑体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 15
        .Bold = False
    End With
    With ActiveDocument.Styles("标题 1").ParagraphFormat
        .SpaceBefore = 24
        .SpaceBeforeAuto = False
        .SpaceAfter = 18
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphCenter
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 1")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub
Sub 中文一级标题()
    With ActiveDocument.Styles("标题 2").Font
        .NameFarEast = "黑体"
        .Name = "黑体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 14
        .Bold = False
    End With
    With ActiveDocument.Styles("标题 2").ParagraphFormat
        .SpaceBefore = 18
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 2")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub 中文二级标题()
    With ActiveDocument.Styles("标题 3").Font
        .NameFarEast = "黑体"
        .Name = "黑体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 14
        .Bold = False
    End With
    With ActiveDocument.Styles("标题 3").ParagraphFormat
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 3")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub 中文三级标题()
    With ActiveDocument.Styles("标题 4").Font
        .NameFarEast = "黑体"
        .Name = "黑体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 12
        .Bold = False
    End With
    With ActiveDocument.Styles("标题 4").ParagraphFormat
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 4")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub 英文章标题()
    With ActiveDocument.Styles("标题 1").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 15
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 1").ParagraphFormat
        .SpaceBefore = 24
        .SpaceBeforeAuto = False
        .SpaceAfter = 18
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphCenter
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 1")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub 英文一级标题()
    With ActiveDocument.Styles("标题 2").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 14
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 2").ParagraphFormat
        .SpaceBefore = 18
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 2")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub


Sub 英文二级标题()
    With ActiveDocument.Styles("标题 3").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 14
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 3").ParagraphFormat
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 3")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub


Sub 英文三级标题()
    With ActiveDocument.Styles("标题 4").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 12
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 4").ParagraphFormat
        .SpaceBefore = 12
        .SpaceBeforeAuto = False
        .SpaceAfter = 6
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    With Selection
     ' Collapse current selection to insertion point.
     .Collapse
     ' Turn extend mode on.
     .Extend
     ' Extend selection to word.
     .Extend
     ' Extend selection to sentence.
     .Extend
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 4")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "标题 1" Or Selection.Style.NameLocal = "标题 2" Or Selection.Style.NameLocal = "标题 3" Or Selection.Style.NameLocal = "标题 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub


Sub 摘要()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Style = ActiveDocument.Styles("正文")
    Selection.Font.Name = "宋体"
    Selection.Font.NameAscii = "Times New Roman"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = False
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
    Selection.EndKey
    Selection.TypeText Chr(13)
End Sub

Sub 目录()
    Set myRange = ActiveDocument.Range(Start:=0, End:=0)
    ActiveDocument.TablesOfContents.Add Range:=myRange, _
    UseFields:=False, UseHeadingStyles:=True, _
    LowerHeadingLevel:=4, _
    UpperHeadingLevel:=1, _
    UseHyperlinks:=True
End Sub


Sub 创建文档还原点()
    On Error GoTo msg
    ActiveDocument.Save
    FName = ActiveDocument.Name
    strs = Split(FName, ".")
    For i = LBound(strs, 1) To (UBound(strs, 1) - 1)
        myname = myname & strs(i) & "."
    Next i
    endformat = strs(UBound(strs, 1))
    timenow = Format(Now, "(还原点yyyy-mm-dd_hh'mm'ss)")
    savename = timenow & myname & endformat
    fpath = ActiveDocument.Path
    ActiveDocument.SaveAs2 fpath & "\" & savename
    ActiveDocument.Close
Documents.Open (fpath & "\" & FName)
    MsgBox "完成！" & Chr(13) & "还原点位于该文档所在文件夹"
    Exit Sub
msg:
    MsgBox "出了点问题，请检查后重试  >_<", Title:="Error", buttons:=vbCritical
End Sub

