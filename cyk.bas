Attribute VB_Name = "cyk"
Sub ȫ�ı����ӱ���()
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
        
        For Each C In Selection.Characters
            If VBA.Asc(C) >= 0 And C.Font.Name <> "Times New Roman" Then
                C.Font.Name = "Times New Roman"
            End If
        Next
        
        End If
        
    Next i
    MsgBox "������б�������"
End Sub

