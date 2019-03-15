Attribute VB_Name = "Module2"
Sub Format_header()
Attribute Format_header.VB_Description = "Format searsport Headers\n"
Attribute Format_header.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format_header Macro
' Format searsport Headers
'


'
    Rows("1:1").Select
    Selection.RowHeight = 20
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
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Rows("2:2").Select
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
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.RowHeight = 93
End Sub
Sub Increment()
'
' Increment Macro
' Auto Increment Cell Values
'

Dim I As Integer

I = 0
Do While ActiveCell.Value <> ""

If ActiveCell.Value <> "" Then
    ActiveCell.Value = I
    I = I + 1
    ActiveCell.Offset(1, 0).Select
        

End If

Loop


End Sub


Sub Button1_Click()
StructuredTextTool.Show
End Sub

