Dim R_Character As Range

    Dim FontSize(5)
    ' 字体大小在5个值之间进行波动，改成需要的大小
    FontSize(1) = "15"
    FontSize(2) = "15.5"
    FontSize(3) = "16"
    FontSize(4) = "16.5"
    FontSize(5) = "16"

    Dim FontName(3)
    '字体名称在三种字体之间进行波动，改成需要的字体，但需要保证系统拥有下列字体
    FontName(1) = "微软雅黑"
    FontName(2) = "微软雅黑"
    FontName(3) = "微软雅黑"


    Dim ParagraphSpace(5)
    '行间距 在一定以下值中均等分布，改成需要的字号
    ParagraphSpace(1) = "12"
    ParagraphSpace(2) = "13"
    ParagraphSpace(3) = "14"
    ParagraphSpace(4) = "12"
    ParagraphSpace(5) = "13"

    'a数值越大，行距波动越大
    a = 1.5
    
    Dim b
    'b数值越大,字距波动越大
    b = 3
    
    
    '不懂原理的话，不建议修改下列代码

    For Each R_Character In ActiveDocument.Characters

        VBA.Randomize
        
        R_Character.Font.Name = FontName(Int(VBA.Rnd * 3) + 1)

        R_Character.Font.Size = FontSize(Int(VBA.Rnd * 5) + 1)

        R_Character.Font.Position = Int(VBA.Rnd * a) + 1

        R_Character.Font.Spacing = Int(VBA.Rnd * b) - 1.3


    Next

    Application.ScreenUpdating = True



    For Each Cur_Paragraph In ActiveDocument.Paragraphs

        Cur_Paragraph.LineSpacing = ParagraphSpace(Int(VBA.Rnd * 5) + 1)


    Next
        Application.ScreenUpdating = True
