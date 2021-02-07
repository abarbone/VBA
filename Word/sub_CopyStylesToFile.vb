Sub CopyStylesToFile(SourceFile As String, DestinationFile As String)
' This function copies all Styles from one file into another file.
' Note that when Styles are imported, the properties of any Styles of the same name within the
' target file will be overwritten to match the properties in the source file. This behavior is
' helpful if you need to update or restore Styles from a central source file.
' Commit date: 2021-02-07

'0.0. Setup.

    '0.1. Set error behavior.

    On Error GoTo ErrorCatch

    '0.2. Initialize variables.
    
    Dim nerr As Integer ' a count of the number of errors.
    
    '0.3. Populate variables.

    ErrorMessage = ""
    nerr = 0

'1.0. Check input.

    '1.1. Check that the Source file path can be found.
    If Dir(SourceFile) = "" Then
        nerr = nerr + 1
        ErrorMessage = ErrorMessage & "Source File not found." & vbCrLf
    End If
    
    '1.2. Check that the Source file path is a docx file.
    If Right(SourceFile, 5) <> ".docx" Then
        nerr = nerr + 1
        ErrorMessage = ErrorMessage & "Source File must have a docx file extension." & vbCrLf
    End If
    
    '1.3. Check Destination file path.
    If Dir(DestinationFile) = "" Then
        nerr = nerr + 1
        ErrorMessage = ErrorMessage & "Destination File not found." & vbCrLf
    End If
    
    '1.4. Check that the Destination file path is a docx file.
    If Right(DestinationFile, 5) <> ".docx" Then
        nerr = nerr + 1
        ErrorMessage = ErrorMessage & "Destination File must have a docx file extension." & vbCrLf
    End If
    
    '1.5. Report error if either path was not found. Otherwise clear nerr and reset Error Message.
    If nerr > 0 Then
        GoTo ErrorCatch
    Else
        nerr = 0
        ErrorMessage = "An unknown error has occurred."
    End If

'2.0. Copy Styles from Source File to Destination File.
    
    '2.1. Open Destination file.
    'Note: This document should automatically become the Active Document.
    Documents.Open FileName:=DestinationFile
        
    '2.2. Copy Styles from Source File.
    ActiveDocument.CopyStylesFromTemplate (SourceFile)
        
    '2.3. Save and Close Destination File.
    ActiveDocument.Close SaveChanges:=wdSaveChanges

'3.0. Close out.

Exit Sub

ErrorCatch:
MsgBox ErrorMessage, vbCritical, "An Error has Occurred"
End

End Sub