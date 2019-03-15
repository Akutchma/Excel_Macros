Attribute VB_Name = "Module6"
Public Sub ErrorCheck()
Attribute ErrorCheck.VB_ProcData.VB_Invoke_Func = " \n14"
'
' GoToSpecial Macro
'
Dim ws As Worksheet, r As range
DebugOutput.Show


For Each ws In ActiveWorkbook.Worksheets
    For Each r In ws.UsedRange
        If IsError(r.Value) Then
            DebugOutput.UpdateProgress "SHEET NAME:  " & ws.Name & "    ADDRESS: " & r.Address & "      THE FORMULA IS: " & r.Formula
        End If
    Next
    For Each r In ws.UsedRange
        If r.Formula Like "*.xl*" Then
            DebugOutput.UpdateProgress "LINK EXISTS IN: " & "SHEET NAME:  " & ws.Name & "    ADDRESS: " & r.Address & "      THE FORMULA IS: " & r.Formula
        End If
    Next
Next


'
End Sub
