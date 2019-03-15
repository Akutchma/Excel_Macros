Attribute VB_Name = "Module3"
Sub Clear_all_Formula()
Attribute Clear_all_Formula.VB_Description = "Clear formulas from ALL cells"
Attribute Clear_all_Formula.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Clear_all_Formula Macro
' Clear formulas from ALL cells
'

'
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    range("A1:G1").Select
End Sub
Sub Unhide_All()
Attribute Unhide_All.VB_Description = "Unhide all cells"
Attribute Unhide_All.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Unhide_All Macro
' Unhide all cells
'

'
    Cells.Select
    Selection.EntireRow.Hidden = False
    Selection.EntireColumn.Hidden = False
End Sub
