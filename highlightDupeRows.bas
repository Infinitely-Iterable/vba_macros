Attribute VB_Name = "highlightDupeRows"
Sub HighlightDuplicateRows()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim lastRow As Long
    Dim dict As Object
    Dim key As String
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A2:C" & lastRow) ' Adjust columns to check
    
    rng.Interior.ColorIndex = xlNone ' Clear any previous highlights
    
    For Each cell In rng.Rows
        key = Join(Application.Transpose(Application.Transpose(cell.Value)), "|")
        If dict.exists(key) Then
            cell.Interior.Color = vbYellow
        Else
            dict.Add key, 1
        End If
    Next cell
End Sub

