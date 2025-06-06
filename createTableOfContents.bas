Attribute VB_Name = "createTableOfContents"
Sub CreateTOC()
    Dim ws As Worksheet, toc As Worksheet
    Dim i As Integer
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Table of Contents").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set toc = Sheets.Add(Before:=Sheets(1))
    toc.Name = "Table of Contents"
    
    toc.Range("A1").Value = "Sheet Index"
    i = 2
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> toc.Name Then
            toc.Hyperlinks.Add Anchor:=toc.Cells(i, 1), _
                Address:="", SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            i = i + 1
        End If
    Next ws
    
    toc.Columns("A").AutoFit
    MsgBox "TOC created!", vbInformation
End Sub


