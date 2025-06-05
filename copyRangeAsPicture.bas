Attribute VB_Name = "copyRangeAsPicture"
Sub copyRangeAsPicture()
    Dim rng As Range
    Set rng = Application.InputBox("Select range to copy as image:", , , , , , , 8)
    If rng Is Nothing Then Exit Sub
    rng.CopyPicture xlScreen, xlPicture
    MsgBox "Range has been copied", vbInformation
End Sub

