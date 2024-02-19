 On Error Resume Next
    
    If Err.Number <> 0 Then
        MsgBox "twój tekst" & Err.Description, vbCritical
        If MsgBox("Czy chcesz kontynuować?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Err.Clear
    End If
