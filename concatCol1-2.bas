Attribute VB_Name = "concat_1_2"
Sub InsertConcatColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the active worksheet or specify a worksheet (e.g., ThisWorkbook.Worksheets("Sheet1"))
    Set ws = ActiveSheet
    
    ' Insert a new column after column B (making it column C)
    ws.Columns("C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Determine the last used row in column A (assuming data exists in column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row and concatenate values from column A and B into the new column C
    For i = 1 To lastRow
        ws.Cells(i, "C").Value = ws.Cells(i, "A").Value & " " & ws.Cells(i, "B").Value
    Next i
End Sub

