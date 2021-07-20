Sub ConvertToGeneral()
    ' Changes the currently-selected text to the format "General".
    ' Commit date: 2021-02-07

    On Error Resume Next

    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With

End Sub
