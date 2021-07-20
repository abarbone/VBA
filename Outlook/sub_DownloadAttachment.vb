'' Downloads an outlook attachement to the Documents\outlook-emails\ folder
''
'' The `outlook-folder' will be created in the user's Document directory
''  via the USERPROFILE environment variable.  The subdirectory is created if it
''  does not already exist.
''
'' The attacment is saved as the name SubjectLine of the email with the
''  attachment name with a "____" separating the names.  Special characters
''  in the subject line are replaced with "_".

Public Sub DownloadAttachment(Mail As Outlook.MailItem)

  Dim SaveFolder As String
  Dim SubjectLine As String
  Dim Attach As Outlook.Attachment
  Dim FileName As String
  Dim FilePath As String
  Dim bad_char As Variant
  Dim env As String

  ' Change to preferred location'
  env = CStr(Environ("USERPROFILE"))
  SaveFolder = env & "\Documents\outlook-emails\"
  SubjectLine = Mail.Subject

  '' Check if folder exists
  SaveFolderExists = Dir(SaveFolder)
  If SaveFolderExists == "" Then
    MkDir SaveFolder
  End If

  ' Remove bad characters from subject line'
  Const BadCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?,/,:"
  For Each bad_char In Split(BadCharacters, ",")
    SubjectLine = Replace(SubjectLine, bad_char, "_")
  Next

  For Each Attach In Mail.Attachments
    FileName = SubjectLine & "___" & Attach.DisplayName
    FilePath = SaveFolder & FileName
    Attach.SaveAsFile FilePath
  Next

End Sub
