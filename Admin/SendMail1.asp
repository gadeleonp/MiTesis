<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Set oMail = Server.CreateObject ("CDONTS.NewMail")

oMail.BodyFormat = 0
oMail.MailFormat = 0

Sender = correoorigen  'Tu email ="Mensajes@Uvirtual.com"
Recipient = email  'Email de destino

  omail.AttachFile(path_arch)

oMail.Value("Reply-To") = session("email")

oMail.From=Sender
oMail.To=Sender
oMail.Bcc=Recipient
oMail.Subject=titulo
oMail.Body=Texto
oMail.Send

Set oMail = Nothing
%>

