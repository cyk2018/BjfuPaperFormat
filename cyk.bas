Attribute VB_Name = "cyk"
Sub 中文全文表格设置()
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        '选中表格
        tb.Select
        '设置表格的格式
        If tb.Title = "" Then
        
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
        tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '表题采用宋体五号
        Selection.Font.Name = "宋体"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
        Selection.Font.Bold = False
                       
        End If
        
    Next i
    MsgBox "请手动输入标题"
End Sub
Sub 英文全文表格设置()
    Dim tbs As Tables, tb As Table
    Set tbs = ActiveDocument.Tables
    n = tbs.Count
    For i = 1 To n
        Set tb = tbs(i)
        '选中表格
        tb.Select
        '设置表格的格式
        If tb.Title = "" Then
        
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
        tb.Title = tabletitle
        Selection.MoveUp unit:=wdParagraph, Count:=1, Extend:=wdExtend
        Selection.ClearFormatting
        
        '表题采用宋体五号
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 10.5
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Selection.Font.Bold = False
        
        Selection.ParagraphFormat.SpaceBefore = 12
        Selection.ParagraphFormat.SpaceAfter = 6
                    
        End If
        
    Next i
    MsgBox "Please enter the title manually"
End Sub

Sub 中文全文图片设置()
    Dim pics As InlineShapes, pic As InlineShape
    Set pics = ActiveDocument.InlineShapes
    n = pics.Count
    For i = 1 To n
        Set pic = pics(i)
        pic.Select
        If pic.Title = "" Then
            Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveRight
            pictitle = Chr(13) & "图 x-x" & " 请输入标题"
            Selection.TypeText pictitle
            pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "宋体"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
        End If
    Next i
    MsgBox "请手动输入标题"
End Sub
Sub 英文全文图片设置()
    Dim pics As InlineShapes, pic As InlineShape
    Set pics = ActiveDocument.InlineShapes
    n = pics.Count
    For i = 1 To n
        Set pic = pics(i)
        pic.Select
        If pic.Title = "" Then
            Selection.MoveRight Count:=1
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
            'Selection.MoveDown unit:=wdParagraph, Count:=1, Extend:=wdExtend
            pictitle = Chr(13) & "Figure x-x" & " Please input the title"
            Selection.TypeText pictitle
            pic.Title = pictitle
            Selection.ClearFormatting
            Selection.Font.Name = "Times New Roman"
            Selection.Font.Size = 10.5
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            Selection.ParagraphFormat.SpaceBefore = 6
            Selection.ParagraphFormat.SpaceAfter = 12
        End If
    Next i
    MsgBox "Please enter the title manually"
End Sub
