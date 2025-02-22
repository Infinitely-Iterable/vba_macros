Attribute VB_Name = "deleteRowsBasedOnFirstCol"
Sub DeleteEmptyRowsBasedOnFirstColumn()
    Dim rng As Range
    Dim i As Long
    
    ' Use the current selection as the range to process.
    Set rng = Selection
    
    ' Loop through the rows from bottom to top.
    For i = rng.Rows.Count To 1 Step -1
        ' Check if the first cell in the row is empty (after trimming spaces).
        If Trim(rng.Cells(i, 1).Value & "") = "" Then
            ' Delete the entire row.
            rng.Rows(i).EntireRow.Delete
        End If
    Next i
End Sub

