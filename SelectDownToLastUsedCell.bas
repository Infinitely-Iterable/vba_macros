Attribute VB_Name = "Module1"
'===========================================================
' Macro: SelectDownToLastUsedCell
' Purpose:
'   Starting from the *single* cell the user has selected,
'   extend the selection down to the last non-blank cell
'   in that same column.
'
' How to use:
'   1. Place this code in a standard module (Alt + F11 ?
'      Insert ? Module).
'   2. Select exactly one cell in any worksheet.
'   3. Run “SelectDownToLastUsedCell”.
'===========================================================

Sub SelectDownToLastUsedCell()
    Dim startCell As Range
    Dim lastRow As Long
    
    '----- Safety check: ensure a single-cell selection -----
    If Selection.CountLarge <> 1 Then
        MsgBox "Please select exactly one cell, then re-run the macro.", _
               vbExclamation, "Selection Error"
        Exit Sub
    End If
    
    Set startCell = ActiveCell
    
    '----- Find the last used row in the same column -----
    lastRow = Cells(Rows.Count, startCell.Column).End(xlUp).Row
    
    ' If the selected cell is below existing data, stop gracefully
    If lastRow < startCell.Row Then
        MsgBox "No data found below the selected cell in this column.", _
               vbInformation, "Nothing to Select"
        Exit Sub
    End If
    
    '----- Select from the starting cell down to the last used cell -----
    Range(startCell, Cells(lastRow, startCell.Column)).Select
End Sub


