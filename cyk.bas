Attribute VB_Name = "cyk"
Sub ȫ�ı������()
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        'ѡ�б��
        tb.Select
        '���ñ��ĸ�ʽ
        If tb.Title = "" Then
        
        '��������߿�
        tb.Borders.InsideLineStyle = wdLineStyleSingle
        tb.Borders.InsideLineWidth = wdLineWidth050pt
        tb.Borders.OutsideLineStyle = wdLineStyleSingle
        tb.Borders.OutsideLineWidth = wdLineWidth050pt
        
        tb.Style.Font.Name = "����"
        tb.Style.Font.Size = 10.5
        tb.Style.Font.Name = "Times New Roman"
        
        
        '��������
        
        Selection.MoveLeft Count:=2
        tabletitle = Chr(13) & "��" & "x-x ���������"
        Selection.TypeText tabletitle
        tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '��������������
        Selection.Font.Name = "����"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Bold = False
                       
        End If
        
    Next i
    MsgBox "������б�������"
End Sub

Sub ȫ��ͼƬ����()
    Dim pics As InlineShapes, pic As InlineShape
    Set pics = ActiveDocument.InlineShapes
    n = pics.Count
    For i = 1 To n
        Set pic = pics(i)
        pic.Select
        If pic.Title = "" Then
            Selection.MoveRight Count:=2
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveRight
            pictitle = Chr(13) & "ͼx-x" & " ���������"
            Selection.TypeText pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "����"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
End Sub
