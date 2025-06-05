Attribute VB_Name = "exportSelectedSheetAsPDF"
Sub ExportSelectedSheetsToPDF()
    Dim sh As Worksheet, fldr As String
    fldr = ThisWorkbook.Path & "\PDF Export\"
    If Dir(fldr, vbDirectory) = "" Then MkDir fldr
    For Each sh In ActiveWindow.SelectedSheets
        sh.ExportAsFixedFormat xlTypePDF, fldr & sh.Name & ".pdf"
    Next sh
    MsgBox "Done" & fldr, vbInformation
End Sub

