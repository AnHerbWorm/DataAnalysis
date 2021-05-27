Attribute VB_Name = "modEmail"
'#######################################################################################
' modEmail
'
' VBA procedures related to creating and sending emails from outlook
'
'#######################################################################################

Option Explicit

Public Sub CreateEmailWithAttachments( _
    ByVal SubjectLine As String, _
    ByVal RecipientsTo As String, _
    Optional ByRef BodyMessage As Range, _
    Optional ByVal RecipientsCC As String, _
    Optional ByVal RecipientsBCC As String, _
    Optional AttachFilePaths As Variant _
)
'#######################################################################################
' CreateEmailWithAttachments
'
' Opens a new email from outlook, with subject line, recipients, and attachments filled.
' The body of the email attachment is read from a range object to the clipboard. This
' method is used because assigning the body via vba will cause the user's default
' email signatures to not be included.
'
' Args:
'   SubjectLine: string
'       Sets the email's subject line.
'   RecipientsTo: string
'       Semi-colon separated list of email addresses. Also accepts written names for
'       existing outlook contacts.
'   BodyMessage: range
'       Range of cell(s) containing a message to be copied to the clipboard. Formatting
'       is preserved.
'   RecipientsCC: string
'       Semi-colon separated list of email addresses. Also accepts written names for
'       existing outlook contacts.
'   RecipientsBCC: string
'       Semi-colon separated list of email addresses. Also accepts written names for
'       existing outlook contacts.
'   AttachFilePaths: string | array
'       FilePath(s) to attach to the email. Can be passed as a string literal or array
'       of strings.
'
' Affects:
'   Nothing in ThisWorkbook. Opens a new email window in the user's outlook with 
'   subject line, all recipient fields, and attachments added. Also puts the email body
'   on clipboard for pasting.
'
' Raises:
'   Nothing. The sub assumes an open instance of outlook.
'#######################################################################################
If Not IsMissing(AttachFilePaths) Then
    Dim iterFile As Variant
    Dim Attachments() As Variant
    
    Dim fsObj As Object
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    
    If VarType(AttachFilePaths) = vbString Then
        Attachments = Array(AttachFilePaths)
    Else
        Attachments = AttachFilePaths
    End If
    
    For Each iterFile In Attachments
        If Not fsObj.FileExists(iterFile) Then
            MsgBox iterFile & " was not found, please check that it was created"
            Exit Sub
        End If
    Next iterFile
End If

Dim outlookApp As Object
Set outlookApp = CreateObject("Outlook.Application")

Dim email As Object
Set email = outlookApp.createitem(0)

If IsMissing(RecipientsCC) Then RecipientsCC = ""
If IsMissing(RecipientsBCC) Then RecipientsBCC = ""

With email
    .display
    .Subject = SubjectLine
    .To = RecipientsTo
    .cc = RecipientsCC
    .bcc = RecipientsBCC
    If Not IsMissing(AttachFilePaths) Then
        For Each iterFile In Attachments
            .Attachments.Add iterFile
        Next iterFile
    End If
End With

If Not IsMissing(BodyMessage) Then
    BodyMessage.Copy
End If

End Sub
