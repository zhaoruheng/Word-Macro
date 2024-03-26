Sub 随机模仿手写()
'
' 随机模仿手写 宏
'
Dim R_Character As Range

    ' 字体大小在下列值之间进行波动，改成需要的大小，重复出现的次数越多，相应出现的概率越大，最小精度0.5
    Dim FontSize() As String
    FontSize = Split("18.5,18.5,18.5,19,18", ",")

    '字体名称在下列字体之间进行波动，改成需要的字体，但需要保证系统拥有下列字体，可以在word查看字体名字
    '请注意，这里的值只影响中文，英文和数字的字体是固定的，如果需要修改英文和数字的字体，可以在下面的代码中修改
    Dim FontName() As String
    FontName = Split("萌妹子体,张维镜手写楷书,手写大象体,陈静的字完整版,汉仪晨妹子W", ",")
    
    ' 推荐字体
    ' "萌妹子体,张维镜手写楷书,手写大象体,陈静的字完整版,汉仪晨妹子W"
    ' 不太理想但可以凑合的字体
    ' "汉仪平安行粗简", "Aa一见钟情 (非商业使用)", "李国夫手写字体"

    'a数值越大，行距越大，波动范围a+x, x∈[-1~1]
    a = 1
    
    'b数值越大，字距越大，波动范围b+x, x∈[-1~1]
    b = 0

    '行间距 在一定以下值中均等分布，改成需要的大小，范围c+x, x∈[0~5]
    c = 25
    
    '不懂原理的话，不建议修改下列代码
    For Each R_Character In ActiveDocument.Characters

        randomlnteger = Int((100 - 1 + 1) * VBA.Rnd + 1)
        If randomlnteger Mod 13 = 0 Then  '13控制倾斜概率，数字越大概率越低
        R_Character.Font.Italic = wdToggle ' 随机倾斜
        ElseIf randomlnteger Mod 27 = 0 Then  '13控制粗体概率，数字越大概率越低
        R_Character.Font.Bold = wdToggle ' 随机粗体
        End If

        VBA.Randomize
        
        ' 数组长度
        FontNameLength = UBound(FontName) - LBound(FontName) 
        FontSizeLength = UBound(FontSize) - LBound(FontSize) 

        ' 字体类型
        R_Character.Font.Name = FontName(Int(VBA.Rnd * FontNameLength) + 1)
        ' 字号大小
        R_Character.Font.Size = FontSize(Int(VBA.Rnd * FontSizeLength) + 1)
        ' 字的上下偏移
        R_Character.Font.Position = Choose(Int(VBA.Rnd * 5) + 1, -1, -0.5, 0, 0.5, 1) + a
        ' 字的左右间距
        R_Character.Font.Spacing = Choose(Int(VBA.Rnd * 5) + 1, -1, -0.5, 0, 0.5, 1) + b
        
        '这是修改字符字体的代码，如果需要修改英文和数字的字体，可以在这里修改
        If R_Character = "。" Or R_Character = "，" Or R_Character = "," Or R_Character = "；" Or R_Character = "’" Or R_Character = "‘" Or R_Character = "“" Or R_Character = "”" Or R_Character = "！" Or R_Character = "？" Or R_Character = "、" Or R_Character = "：" Then
            ' 中文常用标点符号
            R_Character.Font.Name = "汉仪晨妹子W"
        ElseIf Asc(R_Character) >= 48 And Asc(R_Character) <= 57 Then
            ' 数字
            R_Character.Font.Name = "萌妹子体"
        ElseIf Asc(R_Character) >= 97 And Asc(R_Character) <= 122 Or Asc(R_Character) >= 65 And Asc(R_Character) <= 90 Or R_Character = "." Or R_Character = "（" Or R_Character = "）" Or R_Character = "(" Or R_Character = ")" Then
            ' 大小写字母
            R_Character.Font.Name = "汉仪晨妹子W"
        End If

    Next

    For Each Cur_Paragraph In ActiveDocument.Paragraphs
        ' 设置行间距类型为固定值
        Cur_Paragraph.LineSpacingRule = wdLineSpaceExactly
        ' 设置行间距的值
        Cur_Paragraph.LineSpacing = Int(VBA.Rnd * 5) + 1 + c
    Next

		' 设置首行缩进，如不需要注释With到End With这段代码
    With Selection.ParagraphFormat
				' 每个缩进单位长度，厘米
        .FirstLineIndent = CentimetersToPoints(0.35)
				' 设置缩进单位
        .CharacterUnitFirstLineIndent = 2
    End With

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "“"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "”"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Application.ScreenUpdating = True

End Sub