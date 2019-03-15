Attribute VB_Name = "Module4"
Sub Remove_Formats_Formulas_comments()
Attribute Remove_Formats_Formulas_comments.VB_Description = "Remove all formats and formulas and comments"
Attribute Remove_Formats_Formulas_comments.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Remove_Formats_Formulas_comments Macro
' Remove all formats and formulas and comments
'

'
    Cells.Select
    Cells.FormatConditions.Delete
    ActiveWindow.SmallScroll Down:=-18
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveWindow.SmallScroll ToRight:=0
    range("A1:G1").Select
    Application.CutCopyMode = False
End Sub
Sub Delete_Comments()
Attribute Delete_Comments.VB_Description = "Delete Comments in main body"
Attribute Delete_Comments.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Delete_Comments Macro
' Delete Comments in main body
'

'
    range("A1:U347").Select
    Selection.ClearComments
End Sub
