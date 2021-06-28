Sub LookpThroughDocs()
' Loops through and opens all documents at a specific directory. Intended to call other subroutines.
' Commit date: 2021-06-28

'0.0.0. Setup.

    '0.1.0. Set error behavior.

    On Error GoTo ErrorCatch

    '0.2.0. Declare variables.

    Dim file
    Dim Path as String
    Dim FileType As String

    '0.3.0. Populate variables.

    Path = "C:\Users\Alex\Desktop"
    FileType = "*.doc*"

'1.0.0. Check inputs.
' This step not used.

'2.0.0. Perform actions.

    '2.1.0. Start of loop.

    file = Dir(Path & FileType)
    Do While file <> ""
    Documents.Open FileName:=Path & file

    '2.2.0. Call additional subroutines to perfom actions in each file.

    

    '2.3.0. Options for saving or not saving.
    'User Note: Suppress or remove entirely nonpreferred option.

        '2.3.1. Save changes.

        ActiveDocument.Close SaveChanges:=wdSaveChanges

        '2.3.2. Do not save changes.

        ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges

    '2.4.0. Go to next file.
    
    file = Dir()
    Loop

'3.0.0. Close out.

Exit Sub

ErrorCatch:
MsgBox "An unknown error has occcured.", vbCritical, "Error"
End

End Sub
