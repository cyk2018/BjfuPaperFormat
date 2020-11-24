'
'Ver 1.0
'Date 2020.4.23
'Author Silas
'
Public it As Integer
Sub 全文页边距和页眉()
    Selection.WholeStory
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(2.5)
        .BottomMargin = CentimetersToPoints(2.5)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(2)
    End With
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5

    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Delete
    Selection.TypeText text:="北京林业大学本科毕业论文（设计）"
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 9
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Selection.HomeKey
    MsgBox "完成！"
End Sub
'======================================================================

Sub 中文大标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "黑体"
    Selection.Font.Size = 15
    Selection.Font.Bold = False
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

Sub 作者与导师_中文()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 12
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.SpaceBefore = 0
    Selection.ParagraphFormat.SpaceBeforeAuto = False
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.SpaceAfterAuto = False
    Selection.ParagraphFormat.LineUnitBefore = 0
    Selection.ParagraphFormat.LineUnitAfter = 0
    Selection.Font.Bold = False
End Sub

Sub 中文摘要_标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 14
    Selection.Font.Bold = True
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    With Selection.ParagraphFormat
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
    End With
End Sub

Sub 中文摘要_内容()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "楷体"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = False
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
    Selection.EndKey
    Selection.TypeText Chr(13)
End Sub

Sub 中文关键词()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = True
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
    Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
End Sub
'======================================================================

Sub 英文大标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 15
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    Selection.HomeKey unit:=wdLine
    Selection.InsertBreak Type:=wdPageBreak
End Sub

Sub 作者与导师_英文()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 12
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.SpaceBefore = 0
    Selection.ParagraphFormat.SpaceBeforeAuto = False
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.SpaceAfterAuto = False
    Selection.ParagraphFormat.LineUnitBefore = 0
    Selection.ParagraphFormat.LineUnitAfter = 0
    Selection.Font.Bold = False
End Sub

Sub 英文摘要_标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 14
    Selection.Font.Bold = True
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    With Selection.ParagraphFormat
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
    End With
End Sub

Sub 英文摘要_内容()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = False
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
    Selection.EndKey
    Selection.TypeText Chr(13)
End Sub

Sub 英文关键词()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = True
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
    Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
End Sub
'======================================================================

Sub 一级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With ActiveDocument.Styles("标题 1").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "宋体"
        .Size = 14
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 1").ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphCenter
        .LineUnitBefore = 0
        .LineUnitAfter = 1
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 1")
    Selection.HomeKey unit:=wdLine
    Selection.InsertBreak Type:=wdPageBreak
End Sub

Sub 二级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With ActiveDocument.Styles("标题 2").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "宋体"
        .Size = 12
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 2").ParagraphFormat
        .SpaceBefore = 2.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 2")
End Sub

Sub 三级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With ActiveDocument.Styles("标题 3").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "宋体"
        .Size = 10.5
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 3").ParagraphFormat
        .SpaceBefore = 2.5
        .SpaceBeforeAuto = False
        .SpaceAfter = 2.5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphLeft
        .LineUnitBefore = 0.5
        .LineUnitAfter = 0.5
    End With
    Selection.ClearFormatting
    Selection.Style = ActiveDocument.Styles("标题 3")
End Sub

Sub 正文()
    If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定文字！"
        Exit Sub
    End If
    Selection.ClearFormatting
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = False
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 2
        .OutlineLevel = wdOutlineLevelBodyText
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub
'======================================================================

Sub 致谢参考文献附录请使用一级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With ActiveDocument.Styles("标题 1").Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "宋体"
        .Size = 14
        .Bold = True
    End With
    With ActiveDocument.Styles("标题 1").ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 5
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .Alignment = wdAlignParagraphCenter
        .LineUnitBefore = 0
        .LineUnitAfter = 1
    End With
    Selection.Style = ActiveDocument.Styles("标题 1")
    Selection.HomeKey unit:=wdLine
    Selection.InsertBreak Type:=wdPageBreak
End Sub

Sub 参考文献_正文()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 9
    Selection.Font.Bold = False
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .CharacterUnitFirstLineIndent = 0
        .FirstLineIndent = CentimetersToPoints(0)
    End With
End Sub

Sub 表标题_图标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    Selection.ClearFormatting
    Selection.Font.Name = "黑体"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = True
    For Each C In Selection.Characters
        If VBA.Asc(C) >= 0 And C.Font.Name <> "Times New Roman" Then
            C.Font.Name = "Times New Roman"
        End If
    Next
End Sub

Function TableLine()
    With Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth100pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth100pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth100pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
End Function

Sub 表格设置()
    If Selection.Type = wdSelectionIP Or Selection.Type = wdSelectionNormal Then
        MsgBox "请选中表格！"
        Exit Sub
    End If
    temp = TableLine()
    rn = Selection.Tables(1).Rows.Count
    Selection.Font.Name = "宋体"
    Selection.Font.Size = 9
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    temp = TableLine()
    Selection.Tables(1).Rows(2).Select
    Selection.MoveDown unit:=wdLine, Count:=(rn - 2), Extend:=wdLine
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth100pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
    Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
End Sub

Sub 全文表格设置()
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        rn = tb.Rows.Count
        tb.Select
        Selection.Font.Name = "宋体"
        Selection.Font.Size = 9
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
        Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        temp = TableLine()
        tb.Rows(2).Select
        Selection.MoveDown unit:=wdLine, Count:=(rn - 2), Extend:=wdLine
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth100pt
            .DefaultBorderColor = wdColorAutomatic
        End With
        Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        Selection.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
    Next i
    MsgBox "完成"
End Sub

Sub 插入分节符()
    Selection.InsertBreak Type:=wdSectionBreakNextPage
End Sub

Sub 插入分页符()
    Selection.InsertBreak Type:=wdPageBreak
End Sub

Sub 首行缩进两字符()
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
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

Sub 创建文档还原点()
    On Error GoTo msg
    ActiveDocument.Save
    Fname = ActiveDocument.Name
    strs = Split(Fname, ".")
    For i = LBound(strs, 1) To (UBound(strs, 1) - 1)
        myname = myname & strs(i) & "."
    Next i
    endformat = strs(UBound(strs, 1))
    timenow = Format(Now, "(还原点yyyy-mm-dd_hh'mm'ss)")
    savename = timenow & myname & endformat
    fpath = ActiveDocument.Path
    ActiveDocument.SaveAs fpath & "\" & savename
    ActiveDocument.Close
    Documents.Open (fpath & "\" & Fname)
    MsgBox "完成！" & Chr(13) & "还原点位于该文档所在文件夹"
    Exit Sub
msg:
    MsgBox "出了点问题，请检查后重试  >_<", Title:="Error", buttons:=vbCritical
End Sub

Sub 删除全文空行()
    Dim ps As Paragraphs
    Set ps = ActiveDocument.Paragraphs
    For Each p In ps
        If p = Chr(13) Then
            p.Range.Delete
        End If
    Next
    MsgBox "完成！"
End Sub

Sub 删除空格()
    Dim myRange As Range
    If Selection.Type = wdSelectionIP Then
        MsgBox "请选定区域！"
    Else
        Set myRange = Selection.Range
        myRange.Find.Execute FindText:=" ", replacewith:="", Replace:=wdReplaceAll
        MsgBox "空格删除完毕！"
    End If
End Sub

Sub 全选()
    Selection.WholeStory
End Sub

Sub 插入脚注()
    With Selection
        With .FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
            .LayoutColumns = 0
        End With
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
End Sub

Sub 生成页码()
    response = MsgBox("【重要提示】" & Chr(13) & "请确保光标位于第1个一级标题前" & Chr(13) & "再点击[确定]  否则点击[取消]", buttons:=vbOKCancel + vbDefaultButton2)
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
        .Footers(wdHeaderFooterPrimary).PageNumbers.NumberStyle = wdPageNumberStyleUppercaseRoman
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

Sub 全文表格添加标题()
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        tb.Select
        Selection.MoveLeft Count:=2
        Selection.MoveUp unit:=wdParagraph, Count:=2, Extend:=wdExtend
        If Selection.text Like "*表*able*" Then
            Selection.Delete
            Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
            Selection.Delete
        End If
        tb.Select
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "表 " & "x.x 请输入标题" & Chr(13) & "Table " & "x.x Please input title"
        Selection.TypeText tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=2, Extend:=wdExtend
        Selection.ClearFormatting
        Selection.Font.Name = "黑体"
        Selection.Font.Size = 9
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Bold = True
        For Each C In Selection.Characters
            If VBA.Asc(C) >= 0 And C.Font.Name <> "Times New Roman" Then
                C.Font.Name = "Times New Roman"
            End If
        Next
    Next i
    tbs(1).Select
    Selection.MoveLeft Count:=2
    MsgBox "完成！请您手动输入标题内容"
End Sub

Sub 全文图片添加标题()
    Dim pics As InlineShapes, pic As InlineShape
    Set pics = ActiveDocument.InlineShapes
    n = pics.Count
    For i = 1 To n
        Set pic = pics(i)
        pic.Select
        Selection.MoveRight Count:=2
        Selection.MoveDown unit:=wdParagraph, Count:=2, Extend:=wdExtend
        If Selection.text Like "*图*igure*" Then
            Selection.Delete
            Selection.MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
        End If
        pic.Select
        Selection.MoveRight Count:=1
        pictitle = Chr(13) & "图 " & "x.x 请输入标题" & Chr(13) & "Figure " & "x.x Please input title"
        Selection.TypeText pictitle
        Selection.MoveUp unit:=wdParagraph, Count:=2, Extend:=wdExtend
        Selection.ClearFormatting
        Selection.Font.Name = "黑体"
        Selection.Font.Size = 9
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Bold = True
        For Each C In Selection.Characters
            If VBA.Asc(C) >= 0 And C.Font.Name <> "Times New Roman" Then
                C.Font.Name = "Times New Roman"
            End If
        Next
    Next i
    pics(1).Select
    Selection.MoveRight Count:=2
    MsgBox "完成！请您手动输入标题内容"
End Sub

Sub 目录标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With Selection.Font
        .NameFarEast = "黑体"
        .Size = 16
        .Bold = True
        .Italic = False
        .ColorIndex = wdAuto
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
End Sub

Sub 目录一级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Bold = True
        .Underline = False
    End With
    With Selection.ParagraphFormat
        .OutlineLevel = 1
        .LeftIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 0
        .CharacterUnitLeftIndent = 0
    End With
End Sub

Sub 目录二级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Bold = False
        .Underline = False
    End With
    With Selection.ParagraphFormat
        .OutlineLevel = 2
        .LeftIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 0
        .CharacterUnitLeftIndent = 1
    End With
End Sub

Sub 目录三级标题()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "请选定区域！"
        Exit Sub
    End If
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .Size = 10.5
        .Bold = False
        .Underline = False
    End With
    With Selection.ParagraphFormat
        .OutlineLevel = 3
        .LeftIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 0
        .CharacterUnitLeftIndent = 2
    End With
End Sub

Sub 另存为PDF()
    MsgBox "该操作约需要1~30秒" & Chr(13) & "点击【确定】开始执行", Title:="提示"
    Fname = ActiveDocument.Name
    strs = Split(Fname, ".")
    For i = LBound(strs, 1) To (UBound(strs, 1) - 1)
        myname = myname & strs(i) & "."
    Next i
    fpath = ActiveDocument.Path
    ActiveDocument.SaveAs2 fpath & "\" & myname & "pdf", fileformat:=wdFormatPDF
    MsgBox "完成！" & Chr(13) & "pdf文件位于该文档所在文件夹"
End Sub

Sub 正文快捷键()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF2), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.正文"
    FindKey(BuildKeyCode(wdKeyF2)).Execute
End Sub

Sub 首行缩进快捷键()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF3), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.首行缩进两字符"
    FindKey(BuildKeyCode(wdKeyF3)).Execute
End Sub

Sub 使用快捷键()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF2), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.正文"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF3), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.首行缩进两字符"
    txt = "【快捷键信息】" & Chr(13) & "F1  遍历选择  (向导模式中使用)" & Chr(13) & "F2  设置正文格式" & Chr(13) & "F3  首行缩进两字符"
    MsgBox txt, Title:="提示"
End Sub

Sub 遍历段落快捷键()
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF2), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.正文"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF3), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.首行缩进两字符"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF1), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.遍历段落"
    FindKey(BuildKeyCode(wdKeyF1)).Execute
End Sub

Sub 遍历段落()
    If it <> 1 Then
        Selection.HomeKey
        Selection.Expand unit:=wdParagraph
        it = 1
        Exit Sub
    End If
    Selection.Move unit:=wdParagraph
    Selection.Expand unit:=wdParagraph
    If Selection.Type = wdSelectionColumn Or Selection.Type = wdSelectionRow Then
        Selection.Expand unit:=wdTable
    ElseIf Selection.Type <> wdSelectionNormal Then
        Do
            Selection.Move unit:=wdParagraph
            Selection.Expand unit:=wdParagraph
        Loop Until Selection.Type = wdSelectionNormal
    End If
End Sub

Sub 作者()
    txt1 = "【北京林业大学本科毕业论文格式助手1.0】" _
    & Chr(13) & "作者：经济管理学院  金融16级 小林同学" _
    & Chr(13) & "当前版本：1.0" _
    & Chr(13) & "创建日期：2020.4.22" _
    & Chr(13) & "======================================"
    txt2 = "最新版本及安装请查阅：https://github.com/xzy103/BjfuPaperFormat" _
    & Chr(13) & "有问题可联系作者：1342638994@qq.com" _
    & Chr(13) & "或关注作者微信公众号：小林的杂货铺" _
    & Chr(13) & "感谢您的支持！^_^"
    MsgBox txt1 & Chr(13) & txt2
End Sub

Sub 访问网站()
    Dim doc_Hlink As Document
    Set doc_Hlink = Documents.Add
    Dim hypTemp As Hyperlink
    Set hypTemp = doc_Hlink.Hyperlinks.Add(Address:="https://github.com/xzy103/BjfuPaperFormat", Anchor:=Selection.Range)
    hypTemp.Follow NewWindow:=False, AddHistory:=True
    doc_Hlink.Close 0
End Sub

Sub 说明书()
txt0 = "北京林业大学本科毕业论文格式助手说明书" & Chr(13)
txt1 = "作者：经济管理学院 金融16级 小林同学" _
& Chr(13) & "当前版本：1.0" _
& Chr(13) & "创建日期2020.4.22" _
& Chr(13) & "======================================"

txt2 = "【全局设置】" _
& Chr(13) & "全文页边距页眉：一键设置，全文有效，包含设置页边距和页眉两部分" _
& Chr(13) & "生成页码：需要已经设置好各级标题格式，再点击" _
& Chr(13) & "全选：选中全文" _
& Chr(13) & "删除全文空行：对全文有效，常用于对全文进行排版前" _
& Chr(13) & "创建文档还原点：在同一文件目录下生成一个当前文档备份" _
& Chr(13) & "另存为PDF：在同一文件目录下生成一个pdf格式文档" & Chr(13)

txt3 = "【主要标题格式】" _
& Chr(13) & "中文大标题：需要先选中再点击，即论文的标题" _
& Chr(13) & "英文大标题：需要先选中再点击，即论文的英文标题" _
& Chr(13) & "一级标题：需要先选中再点击" _
& Chr(13) & "二级标题：需要先选中再点击" _
& Chr(13) & "三级标题：需要先选中再点击" & Chr(13)

txt4 = "【目录格式】" _
& Chr(13) & "目录标题：需要先选中再点击，即目录的目录二字" _
& Chr(13) & "目录一级标题：需要先选中再点击，即目录中的一级标题条目" _
& Chr(13) & "目录二级标题：需要先选中再点击，即目录中的二级标题条目" _
& Chr(13) & "目录三级标题：需要先选中再点击，即目录中的三级标题条目" & Chr(13)

txt5 = "【摘要与关键词】" _
& Chr(13) & "中文摘要_标题：需要先选中再点击，即中文摘要的摘要二字" _
& Chr(13) & "中文摘要_内容：需要先选中再点击，即中文摘要的内容不含摘要二字" _
& Chr(13) & "中文关键词：需要先选中再点击，即设置中文关键词的格式" _
& Chr(13) & "英文摘要_标题：参考中文" _
& Chr(13) & "英文摘要_内容：参考中文" _
& Chr(13) & "英文关键词：参考中文" & Chr(13)

txt6 = "【常用格式】" _
& Chr(13) & "正文格式：需要先选中再点击，即论文的正文段落，不含图片和表格等元素，快捷键F2" _
& Chr(13) & "致谢/参考文献/附录：需要先选中再点击，这些板块的标题请使用一级标题" _
& Chr(13) & "参考文献_正文：需要先选中再点击" _
& Chr(13) & "作者与导师_中文：需要先选中再点击，位于大标题下" _
& Chr(13) & "作者与导师_英文：参考上一条" & Chr(13)

txt7 = "【表格与图片】" _
& Chr(13) & "全文表格添加标题：不需要选中，对全文有效，完成后需要手动键入标题内容" _
& Chr(13) & "全文图片添加标题：参考上一条" _
& Chr(13) & "全文表格格式设置：不需要选中，将全文的表格设置为符合要求的三线表" _
& Chr(13) & "设置表格格式：需要先选中某一表格再点击，效果同上一条" _
& Chr(13) & "表标题/图表题格式：需要先选中图/表的标题部分再点击" & Chr(13)

txt8 = "【格式处理】" _
& Chr(13) & "首行缩进两字符：选中段落或将光标位于需要更改的段落，再点击，快捷键F3" _
& Chr(13) & "全文英文新罗马字体：不需要选中，对全文有效" _
& Chr(13) & "删除空格：选中区域内的所有空格将会被删除（请勿用于英文段落）" _
& Chr(13) & "插入分页符：先将光标位于需要插入的地方，再点击" _
& Chr(13) & "插入分节符：同上，一般在设置页码的时候可能会用到" & Chr(13)

txt9 = "【关于本插件】" _
& Chr(13) & "作者与版本：查看作者与版本" _
& Chr(13) & "访问网站：访问该插件项目的GitHub仓库网站" _
& Chr(13) & "使用说明书：本文" _
& Chr(13) & "使用快捷键：快捷键信息，如果快捷键失效，请点击这一项" & Chr(13)

txt10 = "【向导模式】" _
& Chr(13) & "参见[向导模式]下的使用说明" _
& Chr(13) & Chr(13) & "======================================" & Chr(13)

txt11 = "最新版本及安装请查阅：https://github.com/xzy103/BjfuPaperFormat" _
& Chr(13) & "有问题可联系作者：1342638994@qq.com" _
& Chr(13) & "或关注作者微信公众号：小林的杂货铺" _
& Chr(13) & "感谢您的支持！^_^"

txts = txt1 & Chr(13) & txt2 & Chr(13) & txt3 & Chr(13) & txt4 & Chr(13) & txt5 & Chr(13) & txt6 & Chr(13) & txt7 & Chr(13) & txt8 & Chr(13) & txt9 & Chr(13) & txt10 & Chr(13) & txt11

Path = ActiveDocument.Path
Dim doc As Document
Set doc = Documents.Add
doc.SaveAs Path & "\北京林业大学本科毕业论文格式助手说明书1.0.docx"

With Selection.PageSetup
    .LineNumbering.Active = False
    .Orientation = wdOrientPortrait
    .TopMargin = CentimetersToPoints(2.54)
    .BottomMargin = CentimetersToPoints(2.54)
    .LeftMargin = CentimetersToPoints(3.17)
    .RightMargin = CentimetersToPoints(3.17)
    .Gutter = CentimetersToPoints(0)
    .HeaderDistance = CentimetersToPoints(1.5)
    .FooterDistance = CentimetersToPoints(1.75)
    .PageWidth = CentimetersToPoints(29.7)
    .PageHeight = CentimetersToPoints(41.99)
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
    .LinesPage = 44
    .LayoutMode = wdLayoutModeLineGrid
End With

Selection.Font.Name = "宋体"
Selection.Font.Size = 18
Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
Selection.Font.Bold = True
Selection.TypeText txt0
Selection.MoveDown
Selection.Font.Name = "宋体"
Selection.Font.Size = 14
Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
Selection.Font.Bold = False
Selection.TypeText txts
Selection.WholeStory
Selection.HomeKey
doc.Save

End Sub

Sub 向导模式说明书()
    txt = "【向导模式使用说明】" _
    & Chr(13) & "1  将光标移到需要开始的地方，通常是全文开头" _
    & Chr(13) & "2  点击向导模式最左边的【执行】图标" _
    & Chr(13) & "3  不断按快捷键F1，遍历选中每个段落" _
    & Chr(13) & "4  每次选中一个部分时，鼠标点击对应格式" _
    & Chr(13) & "5  如遇正文段落，可使用快捷键F2"
    MsgBox txt
End Sub

'=======================================================

'以下代码是草稿
Sub test()
'    On Error GoTo msg
'    n = 2 / 1
'    Exit Sub
'msg:
'    MsgBox "出了点问题，请检查后重试  >_<", Title:="Error", buttons:=vbCritical
'MsgBox Selection.Type
'Selection.Expand Unit:=wdParagraph

'CustomizationContext = NormalTemplate
'KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF1), KeyCategory:=wdKeyCategoryCommand, Command:="Normal.BjfuPaperFormat.遍历段落"
'FindKey(BuildKeyCode(wdKeyF1)).Execute

End Sub





