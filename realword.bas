Attribute VB_Name = "realword"
Sub ����ֽڷ�()
    Selection.InsertBreak Type:=wdSectionBreakNextPage
End Sub
Sub �����ҳ��()
 Selection.InsertBreak Type:=wdPageBreak
End Sub
Sub ����ȫ�ı������()
    response = MsgBox("����Ҫ��ʾ��" & Chr(13) & "ȷ�϶�ȫ�����б�����������" & Chr(13) & "������ 1 x 1 ���", buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
        Exit Sub
    End If
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        'ѡ�б��
        tb.Select
        '���ñ��ĸ�ʽ
        'If tb.Title = "" Then
        
        '��������߿�
        tb.Borders.InsideLineStyle = wdLineStyleSingle
        tb.Borders.InsideLineWidth = wdLineWidth050pt
        tb.Borders.OutsideLineStyle = wdLineStyleSingle
        tb.Borders.OutsideLineWidth = wdLineWidth050pt
        
        tb.Style.Font.Name = "����"
        tb.Style.Font.Size = 10.5
        tb.Style.Font.NameAscii = "Times New Roman"
        
        
        '��������
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "�� " & "x-x ���������"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '��������������
        Selection.Font.Name = "����"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
        Selection.Font.Bold = False
                       
        'End If
        
    Next i
    MsgBox "���ֶ��������"
End Sub
Sub ����ѡ�б������()
    'ѡ�б��˵��������Ϊ5
    If Selection.Type <> 5 Then
        MsgBox "��ǰ��δѡ�б��"
        Exit Sub
    End If
    '��������߿�
        Selection.Borders.InsideLineStyle = wdLineStyleSingle
        Selection.Borders.InsideLineWidth = wdLineWidth050pt
        Selection.Borders.OutsideLineStyle = wdLineStyleSingle
        Selection.Borders.OutsideLineWidth = wdLineWidth050pt
        
        Selection.Style.Font.Name = "����"
        Selection.Style.Font.Size = 10.5
        Selection.Style.Font.NameAscii = "Times New Roman"
        
        
        '��������
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "�� " & "x-x ���������"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '��������������
        Selection.Font.Name = "����"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
        Selection.Font.Bold = False
    MsgBox "���ֶ��������"
End Sub
Sub Ӣ��ȫ�ı������()
    response = MsgBox("����Ҫ��ʾ��" & Chr(13) & "ȷ�϶�ȫ�����б�����������" & Chr(13) & "������ 1 x 1 ���", buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
        Exit Sub
    End If
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        'ѡ�б��
        tb.Select
        '���ñ��ĸ�ʽ
        'If tb.Title = "" Then
        
        '��������߿�
        tb.Borders.InsideLineStyle = wdLineStyleSingle
        tb.Borders.InsideLineWidth = wdLineWidth050pt
        tb.Borders.OutsideLineStyle = wdLineStyleSingle
        tb.Borders.OutsideLineWidth = wdLineWidth050pt
        
        tb.Style.Font.Name = "����"
        tb.Style.Font.Size = 10.5
        tb.Style.Font.NameAscii = "Times New Roman"
        
        
        '��������
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "Table" & " x-x Please input the title"
        Selection.TypeText tabletitle
        'tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '��������������
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
Sub Ӣ��ѡ�б������()
    'ѡ�б��˵��������Ϊ5
    If Selection.Type <> 5 Then
        MsgBox "��ǰ��δѡ�б��"
        Exit Sub
    End If
    '��������߿�
        Selection.Borders.InsideLineStyle = wdLineStyleSingle
        Selection.Borders.InsideLineWidth = wdLineWidth050pt
        Selection.Borders.OutsideLineStyle = wdLineStyleSingle
        Selection.Borders.OutsideLineWidth = wdLineWidth050pt
        
        Selection.Style.Font.Name = "����"
        Selection.Style.Font.Size = 10.5
        Selection.Style.Font.NameAscii = "Times New Roman"
        
        
        '��������
        
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
    MsgBox "���ֶ��������"
End Sub

Sub ����ȫ��ͼƬ����()
    response = MsgBox("����Ҫ��ʾ��" & Chr(13) & "ȷ�϶�ȫ�����б�����������" & Chr(13) & "�����ܰ�����ʽ�е�ͼƬ��", buttons:=vbOKCancel + vbDefaultButton2)
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
            pictitle = Chr(13) & "ͼ x-x" & " ���������"
            Selection.TypeText pictitle
            'pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "����"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
       ' End If
    Next i
    MsgBox "���ֶ��������"
End Sub
Sub ����ѡ��ͼƬ����()
    'ѡ�б��˵��������Ϊ5
    If Selection.Type = wdSelectionShape Or Selection.Type = wdSelectionShape Then
        MsgBox "��ǰ��δѡ��ͼƬ"
        Exit Sub
    End If
    Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveRight
            pictitle = Chr(13) & "ͼ x-x" & " ���������"
            Selection.TypeText pictitle
            'pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "����"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
End Sub


Sub Ӣ��ȫ��ͼƬ����()
    response = MsgBox("����Ҫ��ʾ��" & Chr(13) & "ȷ�϶�ȫ�����б�����������" & Chr(13) & "�����ܰ�����ʽ�е�ͼƬ��", buttons:=vbOKCancel + vbDefaultButton2)
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

Sub Ӣ��ѡ��ͼƬ����()
    'ѡ�б��˵��������Ϊ5
    If Selection.Type = wdSelectionShape Or Selection.Type = wdSelectionShape Then
        MsgBox "��ǰ��δѡ��ͼƬ"
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



Sub ȫ��ҳ�߾��ҳü_ҳ�Ÿ�ʽ()
response = MsgBox("����Ҫ��ʾ��" & Chr(13) & "ȷ�Ϻ�Ḳ�����޸ĵ�ҳü" & Chr(13) & "��������ȷ��", buttons:=vbOKCancel + vbDefaultButton2)
    If response = 1 Then
        response = MsgBox("���ٴ�ȷ�ϡ�" & Chr(13) & "��ȷ���޸���", buttons:=vbOKCancel + vbDefaultButton2)
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
    Selection.TypeText Text:="���ӿƼ���ѧѧʿѧλ���� "
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Font.Name = "����"
    Selection.Font.Size = 10.5
   

    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Delete
    Selection.TypeText Text:=""
    Selection.HomeKey unit:=wdLine
    Selection.EndKey unit:=wdLine, Extend:=wdExtend
    Selection.Font.Name = "����"
    Selection.Font.Size = 10.5
    
    Selection.PageSetup.HeaderDistance = 56.6
    Selection.PageSetup.FooterDistance = 56.6 ' ����ҳ�ŵ�ҳ��ױߵľ���
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
             
                'ҳü
                Set oHead = .Headers(wdHeaderFooterPrimary)
                'ҳ��
                Set oFoot = .Footers(wdHeaderFooterPrimary)
                'ҳüȡ�����ӵ�ǰһ��
                oHead.LinkToPrevious = False
                'ҳ�ŵ�ǰһ��
                oFoot.LinkToPrevious = False
            
            End With
        Next i
    End With
    MsgBox "��ɣ�"
End Sub

Sub Ӣ������_ҳü��ʽ()
     If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "��ѡ�����֣�"
        Exit Sub
    End If
    response = MsgBox("�˲�����������Ӣ������ҳü�ĸ�ʽ" & Chr(13), buttons:=vbOKCancel + vbDefaultButton2)
    If response <> 1 Then
            Exit Sub
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '����ҳü
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 10.5
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub


Sub ȫ��Ӣ��ʹ������������()
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
    MsgBox "��ɣ�"
    Exit Sub
msg:
    MsgBox "���˵����⣬���������  >_<", Title:="Error", buttons:=vbCritical
End Sub

Sub ����_����()
    If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "��ѡ�����֣�"
        Exit Sub
    End If
    Selection.ClearFormatting
    Selection.Font.Name = "����"
    Selection.Font.Size = 12
    Selection.Font.Bold = False
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 20 '�̶��о�20��
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        '.LineSpacingRule = wdLineSpace1pt5
        .CharacterUnitFirstLineIndent = 2 '�����������ַ�
        .OutlineLevel = wdOutlineLevelBodyText
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub
Sub ����_Ӣ��()
    If Selection.Type = wdSelectionIP Then
        Selection.Expand unit:=wdParagraph
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "��ѡ��Ӣ�ģ�"
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
        .CharacterUnitFirstLineIndent = 2 '�����������ַ�
        .OutlineLevel = wdOutlineLevelBodyText
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
End Sub

Sub ����ҳ��()
    response = MsgBox("����Ҫ��ʾ��" & Chr(13) & "��ȷ�����λ�ڵ�һ�µ�һ������ǰ" & Chr(13) & "�ٵ��[ȷ��]  ������[ȡ��]", buttons:=vbOKCancel + vbDefaultButton2)
    If response = 1 Then
        response = MsgBox("���ٴ�ȷ�ϡ�" & Chr(13) & "��ȷ�Ϲ��λ����ȷλ��", buttons:=vbOKCancel + vbDefaultButton2)
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
        .Footers(wdHeaderFooterPrimary).PageNumbers.NumberStyle = wdPageNumberStyleUppercaseRoman '��һ��֮ǰ������������
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
    MsgBox "��ɣ�"
End Sub


Sub �����ע()
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
        .ParagraphFormat.LineSpacingRule = pbLineSpacingSingle '�����о�
        .ParagraphFormat.LineUnitBefore = 0 '��ǰ��
        .ParagraphFormat.LineUnitAfter = 0 '�κ��
        Set myRange = ActiveDocument.Sections(1).Range
    If myRange.Footnotes.NumberingRule = wdRestartSection Then
        myRange.Footnotes.NumberingRule = wdRestartPage
    End If
        .Footnotes.Add Range:=Selection.Range, Reference:=""
    End With
End Sub

Sub �����±���()
    With ActiveDocument.Styles("���� 1").Font
        .NameFarEast = "����"
        .Name = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 15
        .Bold = False
    End With
    With ActiveDocument.Styles("���� 1").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 1")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub
Sub ����һ������()
    With ActiveDocument.Styles("���� 2").Font
        .NameFarEast = "����"
        .Name = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 14
        .Bold = False
    End With
    With ActiveDocument.Styles("���� 2").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 2")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub ���Ķ�������()
    With ActiveDocument.Styles("���� 3").Font
        .NameFarEast = "����"
        .Name = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 14
        .Bold = False
    End With
    With ActiveDocument.Styles("���� 3").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 3")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub ������������()
    With ActiveDocument.Styles("���� 4").Font
        .NameFarEast = "����"
        .Name = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 12
        .Bold = False
    End With
    With ActiveDocument.Styles("���� 4").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 4")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub Ӣ���±���()
    With ActiveDocument.Styles("���� 1").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 15
        .Bold = True
    End With
    With ActiveDocument.Styles("���� 1").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 1")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub

Sub Ӣ��һ������()
    With ActiveDocument.Styles("���� 2").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 14
        .Bold = True
    End With
    With ActiveDocument.Styles("���� 2").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 2")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub


Sub Ӣ�Ķ�������()
    With ActiveDocument.Styles("���� 3").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 14
        .Bold = True
    End With
    With ActiveDocument.Styles("���� 3").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 3")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub


Sub Ӣ����������()
    With ActiveDocument.Styles("���� 4").Font
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Name = "Times New Roman"
        .Size = 12
        .Bold = True
    End With
    With ActiveDocument.Styles("���� 4").ParagraphFormat
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
    Selection.Style = ActiveDocument.Styles("���� 4")
    Selection.GoToPrevious (wdGoToLine)
    If Selection.Style.NameLocal = "���� 1" Or Selection.Style.NameLocal = "���� 2" Or Selection.Style.NameLocal = "���� 3" Or Selection.Style.NameLocal = "���� 4" Then
        Selection.GoToNext (wdGoToLine)
        Selection.ParagraphFormat.SpaceBefore = 0
    Else
        Selection.GoToNext (wdGoToLine)
    End If
End Sub


Sub ժҪ()
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "��ѡ������"
        Exit Sub
    End If
    Selection.Style = ActiveDocument.Styles("����")
    Selection.Font.Name = "����"
    Selection.Font.NameAscii = "Times New Roman"
    Selection.Font.Size = 10.5
    Selection.Font.Bold = False
    Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2
    Selection.EndKey
    Selection.TypeText Chr(13)
End Sub

Sub Ŀ¼()
    Set myRange = ActiveDocument.Range(Start:=0, End:=0)
    ActiveDocument.TablesOfContents.Add Range:=myRange, _
    UseFields:=False, UseHeadingStyles:=True, _
    LowerHeadingLevel:=4, _
    UpperHeadingLevel:=1, _
    UseHyperlinks:=True
End Sub


Sub �����ĵ���ԭ��()
    On Error GoTo msg
    ActiveDocument.Save
    FName = ActiveDocument.Name
    strs = Split(FName, ".")
    For i = LBound(strs, 1) To (UBound(strs, 1) - 1)
        myname = myname & strs(i) & "."
    Next i
    endformat = strs(UBound(strs, 1))
    timenow = Format(Now, "(��ԭ��yyyy-mm-dd_hh'mm'ss)")
    savename = timenow & myname & endformat
    fpath = ActiveDocument.Path
    ActiveDocument.SaveAs2 fpath & "\" & savename
    ActiveDocument.Close
Documents.Open (fpath & "\" & FName)
    MsgBox "��ɣ�" & Chr(13) & "��ԭ��λ�ڸ��ĵ������ļ���"
    Exit Sub
msg:
    MsgBox "���˵����⣬���������  >_<", Title:="Error", buttons:=vbCritical
End Sub

