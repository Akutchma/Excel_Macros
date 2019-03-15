Attribute VB_Name = "Module5"
Sub duplicates()
        'Declare variables to be used
    Dim ColIndex As Integer
    Dim RowIndex As Integer
    
    'Set a column to be used as an index. This will be used to reference what column is selected
    ColIndex = ActiveCell.Column
    
    'Count how many rows are in the workbook.
    RowIndex = Application.ActiveSheet.Cells(Rows.Count, ColIndex).End(xlUp).Row
    
    'Select a range of cells, starting on the second row to negate the effect of grouped cells in the header.
    range(Cells(2, ColIndex), Cells(RowIndex, ColIndex)).Select

    'Format any duplicate values that are found inside of the selection.
        With Selection
            .FormatConditions.Delete
            .FormatConditions.AddUniqueValues
            .FormatConditions(1).DupeUnique = xlDuplicate
            .FormatConditions(1).Interior.Color = RGB(153, 102, 255)
        End With
            
End Sub
