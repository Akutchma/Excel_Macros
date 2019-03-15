Attribute VB_Name = "Module1"
Sub row_height()
Attribute row_height.VB_Description = "Auto set row height"
Attribute row_height.VB_ProcData.VB_Invoke_Func = " \n14"
'
' row_height Macro
' Auto set row height
'

'
    Rows("3:373").Select
    Selection.RowHeight = 20
End Sub
Sub row_format()
Attribute row_format.VB_ProcData.VB_Invoke_Func = " \n14"
'
' row_format Macro
'

'
    Rows("3:360").Select
    Selection.RowHeight = 20
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Consolas"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Consolas"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
