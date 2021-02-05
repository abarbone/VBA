Sub CreateFoldersFromArray(ByRef FolderArray() As String)
' Creates folders based on a 1 dimensional array of folder paths.
' Commit date: 2021-02-05

On Error GoTo ErrorCatch
ErrorMessage = "An unknown error has occurred."

' Declare variables
Dim i As Long
Dim Parent As String

' Loop through array
For i = LBound(FolderArray) To UBound(FolderArray)
    
    ' Check that the Parent Folder can be created.
    Parent = ParentFolder(FolderArray(i))
    If Dir(Parent, vbDirectory) = "" Then
        ErrorMessage = "Cannot find parent folder: " & Parent _
        & vbCrLf & vbCrLf & "Operation aborted."
        GoTo ErrorCatch
    End If
    
    ' Check if folder already exists. If not, create it.
    If Dir(FolderArray(i), vbDirectory) = "" Then
        MkDir FolderArray(i)
    End If
Next i

Exit Sub

ErrorCatch:
MsgBox ErrorMessage, vbCritical, "An Error has Occurred"
End

End Sub