Attribute VB_Name = "quickVersionBackup"
Sub quickVersionBackup()
    Dim f$, ext$, base$
    base = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    ext = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    f = ThisWorkbook.Path & "\" & base & "_" & Format(Now, "yyyy-mm-dd-hhmm") & ext
    ThisWorkbook.SaveCopyAs f
    MsgBox "Saved ? " & f, vbInformation
End Sub

