<%
dim db
Dim objUpload 'stores an Upload control object
Dim objEmail 'stores a Mailer control object
Dim strPath 'stores path to a directory we create at same level as this ASP page 
dim autor
dim opcion

'The SendMail function will throw an exception if the operation is
' unsuccessful, so we enable inline error trapping
On Error Resume Next

Set objUpload = Server.CreateObject("Dundas.Upload") 
Set oMail = Server.CreateObject("CDONTS.NewMail") 
set db = server.CreateObject("SQLComandos.Comandos")

oMail.BodyFormat = 0
oMail.MailFormat = 0

'create a directory at same level as this ASP page
'this directory is used to store the uploaded files (which are renamed with a GUID preceding
'the original filename)
'NOTE: to delete these files you would have to use the ImpersonateUser method of
'the Upload control since the IUSR_ account SHOULD NOT have permission to delete files

objUpload.DirectoryCreate  "\aecys\paecycod\temp" 


'save the uploaded files. THIS POPULATES THE UPLOAD CONTROL'S COLLECTIONS!

autor = objUpload.Form("autor")
Response.Write autor
objUpload.Save strPath




		copia = objUpload.Form("copia")
		'especifica la dirección a la que se enviara una copia
		oMail.Cc = copia


'especifica el tema del mensaje
tema = objUpload.Form("tema")
oMail.Subject = tema


'especifica el remitente del correo
'objEmail.FromAddress = "Correo@aecys.org"
oMail.From = "aecys@ing.usac.edu.gt"

'specify an SMTP Relay server. This increases the speed and reliability of the operation
'objEmail.SMTPRelayServers.Add "SomeSmtpServer.com"

'initialize the HtmlBody property, we'll throw a header into it
contenido = objUpload.Form("contenido")
Response.Write contenido
HTMLBody =  "<Html><Head></Head><Body><H2>" & contenido


For Each Item in objUpload.Files
	'verifica que el campo del archivo sea el valido
	If (Item.TagName = "txtarchivo") Then
		'If InStr(1,Item.ContentType,"audio") Then
			'objEmail.HtmlEmbeddedObjs.Add Item.Path, 1, Item.OriginalPath
			oMail.AttachFile (Item.Path)
			'objEmail.HtmlBody = objEmail.HtmlBody & "<a href=cid:1>Docu</a>"
			HTMLBody = HTMLBody & "<a href=cid:1>Documento</a>"
		'End If
	End If
	If (Item.TagName = "txtImage") Then
		If InStr(1,Item.ContentType,"image") Then
			'objEmail.HtmlEmbeddedObjs.Add Item.Path, 2, Item.OriginalPath
			oMail.AttachFile (Item.Path)
			'objEmail.HtmlBody = objEmail.HtmlBody & "<IMG SRC=cid:2>"
			HTMLBody = HTMLBody & "<IMG SRC=cid:2>"
		End If
	End If
Next

'terminar html body colocando las etiquetas html finales
	'objEmail.HTMLBody = objEmail.HTMLBody & "</body></html>"
		HTMLBody = HTMLBody & "</body></html>"
		oMail.Body = HTMLBody
		
		destino = objUpload.Form("lstus")
		
		sql = "select mail from usuario where id in (" & destino & ") and recibir mail = 1"
		set rs = db.SQLSelect("DBBAECYS",sql)
		Response.Write sql
		Response.End
		
	if not rs.eof then
		while rs.eof <> true
			correo = rs("mail") 
			'especifica el destinatario del mensaje
			response.write correo
			'objEmail.TOs.Add correo
			'objEmail.SendMail
			

			'Prueba El exito/fallo
			If Err.Number <> 0 Then
				Response.Write "No se pudo enviar el mensaje, el siguiente error ocurrio: " & Err.Description
				'Response.Redirect ("previewforo.asp?desc="&"No se pudo enviar el mensaje, el siguiente error ocurrio: " & Err.Description)
			Else
			'Exito
				Response.Write "The correo html fue enviado exitosamente."
				'Response.Redirect ("previewforo.asp?desc="&"The correo html fue enviado exitosamente.")
			End If
		rs.movenext 
		wend 
	end if
	set rs = nothing
'enviar el correo'Response.Write "aqui<br>" 'Response.Write tema'Response.Write contenido'Response.Write objEmail.HTMLBody
Response.Redirect ("previewforo.asp?desc="&objEmail.HTMLBody)
'Response.End 

'Liberar recursos
Set oMail = Nothing
Set objUpload = Nothing

set db = nothing

%>

  
