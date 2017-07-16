Attribute VB_Name = "Mod_Outlook"
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine will BCC the designated email for all emails sent
'
' Requirements:
' Reference -> Microsoft Outlook ##.# Object Library
'
' Inputs:
' varaintWksArray()
' Type: variant
' Desc: string array with the names of the worksheets and worksheet flag
' varaintWksArray(x,1) -> worksheet name
' varaintWksArray(x,2) -> flag to clear contents
'
' Important Info:
' The Sub should be in "application->ItemSend
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' declare
Dim objRecip As Recipient
Dim strMsg As String, strBcc As String
Dim res As Integer

' start
On Error Resume Next
' #### USER OPTIONS ####
' address for Bcc -- must be SMTP address or resolvable
' to a name in the addressbook
strBcc = "????@??????.com"
Set objRecip = Item.Recipients.Add(strBcc)
objRecip.Type = olBCC
 
If Not objRecip.Resolve Then
    strMsg = "Could not resolve the Bcc recipient. " & "Do you want still to send the message?"
    res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, "Could Not Resolve Bcc Recipient")
    If res = vbNo Then
        Cancel = True
    End If
End If

' object cleanup & reset
On Error GoTo 0
Set objRecip = Nothing
End Sub
Sub Outlook_SendEmail(wksEmailAddressSrc As Worksheet, stringSectionId As String, stringFileId As String, stringEmailSubj As String, collGrievences As Collection)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' this subroutine gets an email set of addresses in <wksEmailAddressSrc> w/ the subject line stringEmailSubj
'
' Requirements:
' <function, class, sub>
'
' Inputs:
' wksEmailAdressSrc
' Type: worksheet
' Desc: the worksheet where the email addresses are held
'
' stringSectionId
' Type: string
' Desc: the ID to find the row where the email section starts
'
' stringFileId
' Type: string
' Desc: the ID to determine the row where the emails are held
'
' stringEmailSubj
' Type: string
' Desc: the subject line of the email
'
' collGrievences
' Type: collection
' Desc: the collectioin which holds an array with the grievance information in it
' first reccored in collection is the header
' variantArray(x,2) -> long; column of row in data
' variantArray(1,1) -> string; Member full name
' variantArray(2,1) -> string; Days to Resolve
' variantArray(3,1) -> string; Line of Business (LOB)
' variantArray(4,1) -> string; Plan Code
' variantArray(5,1) -> string; Provider NPI
' variantArray(6,1) -> string; Call Tracking provider ID
' variantArray(7,1) -> string; Call Tracking member ID
'
' Important Info:
' - Must include the library "Microsoft Outlook 14.0 Object Library" under Tools -> Reference, may be a newer version in an updated Excel Application
'
' Return:
' None
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' declare variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Dim outlookApp As Object
Dim outlookEmail As MailItem
Dim longEmailSectionStartRow As Long, longEmailRow As Long, longEmailCount As Long
Dim stringEmailBody As String, stringEmailAddresses As String
 
' loop
Dim a As Long, b As Long, c As Long
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' set objects
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set outlookApp = CreateObject("Outlook.Application")
Set outlookEmail = outlookApp.CreateItem(olMailItem)
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' initialize variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
longEmailSectionStartRow = 0
longEmailRow = 0
longEmailCount = 0
stringEmailBody = Empty
stringEmailAddresses = Empty
 
' loop
a = 1
b = 1
c = 1
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' begin
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

' get rows of the email section
longEmailSectionStartRow = Row_Find(stringSectionId, 1, 1, Row_GetLast(wksEmailAddressSrc, 1)) + 2
longEmailRow = Row_Find(stringFileId, 2, longEmailSectionStartRow, Row_GetLast(wksEmailAddressSrc, 2))
longEmailCount = Column_GetLast(wksEmailAddressSrc, longEmailRow) - 2
 
' set "To" email addresses
For a = 1 To longEmailCount
    stringEmailAddresses = stringEmailAddresses & Cells(longEmailRow, a + 2).Text & ";"
Next a
 
' set email body
For b = 1 To collGrievences.Count
    For c = 1 To UBound(collGrievences.Item(b), 1)
        stringEmailBody = stringEmailBody & CStr(collGrievences.Item(b)(c, 1)) & "; "
    Next c
   
    ' ASCII caraige return
    stringEmailBody = stringEmailBody & Chr(13)
Next b
 
' configure email and send
With outlookEmail
    .To = stringEmailAddresses
    .Subject = stringEmailSubj
    .Body = stringEmailBody
    .Send
End With
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' end
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
' reset objects
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''’’'''

Set outlookApp = Nothing
Set outlookEmail = Nothing
End Sub

