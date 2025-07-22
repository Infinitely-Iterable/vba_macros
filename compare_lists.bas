Attribute VB_Name = "compare_lists"
' This macro is one way meaning it does not list entries in table2 that are not in table1, only the entries in table one that are not in table2.
' You can rename the sheets and tables as necessary. I have other macros in my repository at https://github.com/Infinitely-Iterable/vba_macros that cover other common tasks!
Sub CompareTablesAndListDifferences()

    Dim tbl1 As ListObject
    Dim tbl2 As ListObject
    Dim outputSheet As Worksheet
    Dim r As Long
    Dim c As Long
    Dim found As Boolean
    Dim cellValue1 As Variant
    Dim cellValue2 As Variant
    Dim outputRow As Long
    
    ' Set references to the tables.
    ' Change "Sheet1" and "Table1" to the appropriate sheet name and table name for your first table.
    ' Change "Sheet2" and "Table2" to the appropriate sheet name and table name for your second table.
    Set tbl1 = Worksheets("Sheet1").ListObjects("Table1")
    Set tbl2 = Worksheets("Sheet2").ListObjects("Table2")
    
    Set outputSheet = Worksheets.Add
    outputSheet.Name = "Differences"
    outputRow = 1
    
    For c = 1 To tbl1.ListColumns.Count
        outputSheet.Cells(outputRow, c).Value = tbl1.HeaderRowRange.Cells(1, c).Value
    Next c
    outputRow = outputRow + 1

    For r = 1 To tbl1.ListRows.Count
        found = False
        For r2 = 1 To tbl2.ListRows.Count
            found = True
            For c = 1 To tbl1.ListColumns.Count
                cellValue1 = tbl1.ListRows(r).Range(1, c).Value
                cellValue2 = tbl2.ListRows(r2).Range(1, c).Value
                If cellValue1 <> cellValue2 Then
                    found = False
                    Exit For
                End If
            Next c
            If found Then Exit For
        Next r2
    
        If Not found Then
            For c = 1 To tbl1.ListColumns.Count
                outputSheet.Cells(outputRow, c).Value = tbl1.ListRows(r).Range(1, c).Value
            Next c
            outputRow = outputRow + 1
        End If
    Next r

    MsgBox "Comparison complete! Differences listed in '" & outputSheet.Name & "'."

End Sub
