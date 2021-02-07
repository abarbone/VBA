'' Downloads an outlook email to the Documents\outlook-emails\ folder
''
'' The `outlook-folder' will be created in the user's Document directory
''   via the USERPROFILE environment variable.  The subdirectory is created if
''   it does not already exist.
''
'' Email is saved with the time in which it was received.

Public Sub SaveMessage(Item As Outlook.MailItem)
  Dim sPath As String
  Dim dtDate As Date
  Dim sName As String
  Dim env As String

  env = CStr(Environ("USERPROFILE"))

  sName = Item.Subject
  ReplaceCharsForFileName sName, "_"

  dtDate = Item.ReceivedTime
  sName = Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, _
  vbUseSystem) & Format(dtDate, "-hhnnss", _
    vbUseSystemDayOfWeek, vbUseSystem) & "-" & sName & ".msg"

    '' use My Documents in older Windows.
    sPath = env & "\Documents\outlook-emails\"
    sPathExists = Dir(sPath)

  '' Check if file exists
  If sPathExists == "" Then
    MkDir sPath
  End If

  Debug.Print sPath & sName
  Item.SaveAs sPath & sName, olMSG
End Sub

Private Sub ReplaceCharsForFileName(sName As String, _
  sChr As String _
)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
End Sub
