Function ParentFolder(FullPath As String) As String
    ' This function outputs the parent folder of a file or folder path.
    ' Commit date: 2021-02-07

    ' If there is a trailing backslash, remove it. Necessary for folder paths ending with a backslash.
    If Right(FullPath, 1) = "\" Then
        FullPath = Left(FullPath, Len(FullPath) - 1)
    End If

    ' Break up the Full Path by splitting it up at the forward slashses.
    SplitPath = Split(FullPath, "\")

    ' Rebuild the Path using the split up path, but don't include the final component.
    For i = LBound(SplitPath) To UBound(SplitPath) - 1
        Rebuilt = Rebuilt & SplitPath(i) & "\"
    Next i

    ' Output the rebuilt path
    ParentFolder = Rebuilt

End Function