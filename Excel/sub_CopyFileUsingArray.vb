Sub CopyFileUsingArray(ByRef DestFolderArray() As String, SourceParentFolder As String, SourceFileName As String, Optional Prefix As String = "")
' Copies a file from a single location into the folders specified in a 1 dimensional array of folder paths.
' Commit date: 2021-02-07

'0.0.0. Setup.

    '0.1.0. Set error behavior.

    On Error GoTo ErrorCatch
    ErrorMessage = "An unknown error has occurred."

    '0.2.0. Declare variables.

    Dim i As Long
    Dim DestFileName As String
    Dim DestFullPath As String
    Dim Parent As String

    '0.3.0. Populate variables.

        '0.3.1. Generate Destination File Name
        
        DestFileName = Prefix & SourceFileName

        '0.3.2. Generate Full Source Path.

        If Right(SourceParentFolder, 1) = "\" Then
            SourceFullPath = SourceParentFolder & SourceFileName
        Else
            SourceFullPath = SourceParentFolder & "\" & SourceFileName
        End If

'1.0.0. Check inputs.

    '1.1.0. Check if source file can be found.

    If Dir(SourceFullPath) = "" Then
        ErrorMessage = "Cannot find source file:" & vbCrLf & SourceFullPath _
        & vbCrLf & vbCrLf & "Operation aborted."
        GoTo ErrorCatch
    End If

'2.0.0. Copy files.

    ' Loop through array
    For i = LBound(DestFolderArray) To UBound(DestFolderArray)
    
        ' Create the Full Destination Path.
        If Right(DestFolderArray(i), 1) = "\" Then
            DestFullPath = DestFolderArray(i) & DestFileName
        Else
            DestFullPath = DestFolderArray(i) & "\" & DestFileName
        End If
    
        ' Check that parent folder can be found.
        Parent = ParentFolder(DestFullPath)
        If Dir(Parent, vbDirectory) = "" Then
            ErrorMessage = "Cannot find parent folder:" & vbCrLf & Parent _
            & vbCrLf & vbCrLf & "Operation aborted."
        End If
    
        ' Check if file already exists. If not, copy it.
        If Dir(DestFullPath) = "" Then
            FileCopy SourceFullPath, DestFullPath
        End If
    
    Next i

'3.0.0. Close out.

Exit Sub

ErrorCatch:
MsgBox ErrorMessage, vbCritical, "An Error has Occurred"
End

End Sub