Attribute VB_Name = "InsertCode"
Sub aaa_set_table_and_insert_code()
    ' Usage: Copy RTF from notepad++ and run this Macro
START_INPUT:
    tab_or_space = MsgBox("tab: Yes   space: No", 4, "Indent type")
    If tab_or_space = 6 Then
        indent_char = "^t"
        bk_times = 1
        indent_width = 4
    ElseIf tab_or_space = 7 Then
        space_indent_len = InputBox("space number", "space number", 4)
        indent_char = String(space_indent_len, " ")
        bk_times = space_indent_len
        indent_width = space_indent_len
    End If
        
    font_size = InputBox("font size", "font size", 9)
    line_space = font_size + 3
    
    Set mytable = ActiveDocument.Tables.Add(Selection.Range, 1, 3)
    mytable.Columns(1).SetWidth ColumnWidth:=3 * font_size * 1.2 * 0.5 + 5, _
                                RulerStyle:=wdAdjustProportional
    mytable.Columns(2).SetWidth ColumnWidth:=10, _
                                RulerStyle:=wdAdjustProportional
    With mytable.Cell(1, 3).Range
        .Paste
        .ParagraphFormat.Shading.Texture = wdTextureNone
        .HighlightColorIndex = wdNoHighlight
    End With
    
    ' background RGB=(229,229,229)
    With mytable
        With .Shading
            .Texture = wdTextureNone
            .ForegroundPatternColor = wdColorAutomatic
            .BackgroundPatternColor = 15066597
        End With
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
        ' .AutoFitBehavior (wdAutoFitContent)  '自動條整大小
    End With
    
    
    
    With mytable.Range.ParagraphFormat
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = line_space
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .MirrorIndents = False
        .TextboxTightWrap = wdTightNone
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    
    
    With mytable.Range.Font
        .Name = "Courier New"
        .Size = font_size
    End With

    
    
    
    ' Convert tab to indent
    Dim found As Boolean
    mytable.Cell(1, 3).Range.Select
    With Selection.Find
        .text = "^p" & indent_char
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
    End With
    found = Selection.Find.Execute
    While (found)
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            For i = 1 To bk_times
                Selection.TypeBackspace
            Next
            Selection.Paragraphs.LeftIndent = Selection.Paragraphs.LeftIndent + 1.2 * 0.5 * font_size * indent_width
            mytable.Cell(1, 3).Range.Select
            With Selection.Find
                .text = "^p" & indent_char
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
            End With
            found = Selection.Find.Execute
    Wend
    
        
    Set myRng = mytable.Cell(1, 3).Range
    myRng.End = mytable.Cell(1, 3).Range.End - 1
    
    line_count = myRng.ComputeStatistics(wdStatisticLines)
    
    
    mytable.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    For i = 1 To line_count - 1
        mytable.Cell(1, 1).Range.InsertAfter text:=i
        mytable.Cell(1, 1).Range.InsertParagraphAfter
    Next
    mytable.Cell(1, 1).Range.InsertAfter text:=line_count
End Sub
 
Sub insert_continued_number()
    Line_start = InputBox("Insert start line number", "From", "1")
    Line_end = InputBox("Insert end line number", "To", "50")
    For i = Line_start To Line_end - 1
        Selection.TypeText text:=i
        Selection.TypeParagraph
    Next
    Selection.TypeText text:=Line_end
End Sub


' Ref
' https://blog.csdn.net/code4101/article/details/41802715
' https://superuser.com/questions/1088116/how-to-convert-tabs-to-indents-in-microsoft-word
