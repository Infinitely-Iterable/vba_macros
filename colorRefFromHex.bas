Sub FillCellsWithHexColors()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim hexCode As String
    Dim r As Long, g As Long, b As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' assumes column A has hex codes

    For i = 1 To lastRow
        hexCode = ws.Cells(i, 1).Value ' Column A (change if needed)
        If Left(hexCode, 1) = "#" Then hexCode = Mid(hexCode, 2)
        If Len(hexCode) = 6 Then
            On Error Resume Next
            r = CLng("&H" & Mid(hexCode, 1, 2))
            g = CLng("&H" & Mid(hexCode, 3, 2))
            b = CLng("&H" & Mid(hexCode, 5, 2))
            ws.Cells(i, 2).Interior.Color = RGB(r, g, b) ' Fill color in column B
            On Error GoTo 0
        End If
    Next i
End Sub
