Sub SendEmail()
' Populates an email, including subject, to, cc, and bcc fields.
' Requires that the Reference "Microsoft Outlook 16.0 Object Library" be enabled.
' Commit date: 2021-06-28

'0.0.0. Setup.

    '0.1.0. Set error behavior.

    On Error GoTo ErrorCatch
   
    '0.2.0. Declare variables.

    Dim EmailApp As Outlook.Application
    Dim EmailItem As Outlook.MailItem
    Dim SubjectLine As String
    Dim ToList As String
    Dim CcList As String
    Dim BccList As String
    Dim BodyText As String

    '0.3.0. Populate Variables.

    SubjectLine = "Sample Subject Line"
    ToList = "Employee1@company.com; Employee2@company.com"
    CcList = "Employee3@company.com"
    BccList = ""
    BodyText = "Please send me the thing. Thanks."

'1.0.0. Check inputs.
' Not used.

'2.0.0. Create email.

    '2.1.0. Change to Outlook and create a new message.
    
    Set EmailApp = New Outlook.Application
    Set EmailItem = EmailApp.CreateItem(olMailItem)

    '2.2.0. Populate the email.

    With EmailItem
        .Subject = SubjectLine
        .To = ToList
        .CC = CcList
        .BCC = BccList
        .body = BodyText
        .Display
    End With

    '2.3.0. Reset object variables.

    Set EmailApp = Nothing
    Set EmailItem = Nothing
    
'3.0.0. Close out.

Exit Sub

ErrorCatch:
MsgBox "An error occurred while generating the email.", vbCritical, "Error"
End

End Sub
