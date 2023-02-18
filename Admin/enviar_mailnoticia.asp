<%
Dim objEmail 'stores a Mailer control object
Dim strPath 'stores path to a directory we create at same level as this ASP page 

'The SendMail function will throw an exception if the operation is
' unsuccessful, so we enable inline error trapping
On Error Resume Next


Set objEmail = Server.CreateObject("Dundas.Mailer") 'Mailer object

'create a directory at same level as this ASP page
' this directory is used to store the uploaded files (which are renamed with a GUID preceding
' the original filename)
'NOTE: to delete these files you would have to use the ImpersonateUser method of
' the Upload control since the IUSR_ account SHOULD NOT have permission to delete files




'specify the recipient of this message
objEmail.TOs.Add "aecys@ing.usac.edu.gt"

'specify the subject of the email
objEmail.Subject = "This is the subject."

'specify the sender of the message
objEmail.FromAddress = "MyUsername@MyDomain.com"

'specify an SMTP Relay server. This increases the speed and reliability of the operation
objEmail.SMTPRelayServers.Add "SomeSmtpServer.com"

'initialize the HtmlBody property, we'll throw a header into it
objEmail.HTMLBody = "<Html><Head></Head><Body><H2>This is the body.</H2><BR><BR>"

'finish html body by adding closing html tags
objEmail.HTMLBody = objEmail.HTMLBody & "</body></html>"

'send the email
objEmail.SendMail

'test for success/failure
If Err.Number <> 0 Then
'an error occurred so output error message
Response.Write "Sorry, the following error occurred: " & Err.Description
Else
'success!
Response.Write "The html email was successfully sent."
End If

'release resources
Set objEmail = Nothing
Set objUpload = Nothing
%>
