Attribute VB_Name = "toSqlValue"
Sub ConvertRangeToSQL()
    Dim rng As Range
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long
    Dim cellValue As Variant
    Dim sql As String
    Dim rowSQL As String
    
    ' Use the current selection as the range to convert.
    Set rng = Selection
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count
    sql = ""
    
    ' Loop through each row in the selection.
    For i = 1 To rowCount
        rowSQL = "("
        ' Loop through each cell (column) in the row.
        For j = 1 To colCount
            cellValue = rng.Cells(i, j).Value
            ' Check if cell is numeric and not empty.
            If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
                rowSQL = rowSQL & cellValue
            Else
                ' Wrap text in single quotes and escape any internal quotes.
                rowSQL = rowSQL & "'" & Replace(CStr(cellValue), "'", "''") & "'"
            End If
            ' Append comma between columns, except after the last column.
            If j < colCount Then
                rowSQL = rowSQL & ", "
            End If
        Next j
        rowSQL = rowSQL & ")"
        ' Append a comma and line break if this is not the last row.
        If i < rowCount Then
            rowSQL = rowSQL & "," & vbCrLf
        End If
        sql = sql & rowSQL
    Next i
    
    ' Output the generated SQL statement in a new worksheet.
    Dim ws As Worksheet
    Set ws = Worksheets.Add
    ws.Range("A1").Value = sql
    
    MsgBox "SQL statement generated in new worksheet.", vbInformation
End Sub

