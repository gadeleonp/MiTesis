<%
		Dim objUpload 'stores an Upload control object
		Dim objEmail 'stores a Mailer control object
		Dim strPath 'stores path to a directory we create at same level as this ASP page 
		dim autor
		dim opcion
		dim db

'''''''''
    HTML = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
    HTML = HTML & "<html>"
    HTML = HTML & "<head>" 
    HTML = HTML & "<title>Sending CDONTS Email Using HTML</title>"
    HTML = HTML & "</head>"
    HTML = HTML & "<body bgcolor=""FFFFFF"">"
    HTML = HTML & "<p><font size =""3"" color = ""ccc20001"" face=""Arial""><strong>"
    HTML = HTML & "Name Of Store</strong><br>"
    HTML = HTML & "Incoming Customer Order</strong></p>"
    HTML = HTML & "<p align = ""center"">mi prueba gracias</p>"
    HTML = HTML & "<IMG SRC=cid:2>"
    HTML = HTML & "<a href=cid:1>Documento</a>"
    HTML = HTML & "</body>"
    HTML = HTML & "</html>"
    
''''''''


		'The SendMail function will throw an exception if the operation is
		' unsuccessful, so we enable inline error trapping
		On Error Resume Next

		set db = server.CreateObject("SQLComandos.Comandos")
		Set objUpload = Server.CreateObject("Dundas.Upload.2") 

		'Quitado 
		'Set objEmail = Server.CreateObject("Dundas.Mailer") 

		Set oMail = Server.CreateObject ("CDONTS.NewMail")

		oMail.BodyFormat=0
		oMail.MailFormat=0
		
		'create a directory at same level as this ASP page
		'this directory is used to store the uploaded files (which are renamed with a GUID preceding
		'the original filename)
		'NOTE: to delete these files you would have to use the ImpersonateUser method of
		'the Upload control since the IUSR_ account SHOULD NOT have permission to delete files
		strPath = Server.MapPath(".") & "\temp" 
		objUpload.DirectoryCreate strPath
		destino = ""

		'save the uploaded files. THIS POPULATES THE UPLOAD CONTROL'S COLLECTIONS!
		objUpload.Save strPath

		autor = objUpload.Form("autor")
		opcion = objUpload.Form("opcion")



		dim rs

		  
			if opcion = "1" and not isempty(autor) then
				sql = "select nombre,apellido, mail from usuario where id = "& autor
				
					set rs = db.SQLSelect("DBBAECYS",sql)
					if not rs.eof then
						destino = rs("mail")
						nombre_completo = rs("Nombre") & " " & rs("apellido")
					end if
				set rs = nothing
				set db = nothing
			else
				destino = objUpload.Form("destino")

				copia = objUpload.Form("copia")

			
			end if 

		oMail.Value("Reply-To") = "aecys@ing.usac.edu.gt"
		
		'especifica el destinatario del mensaje
		'QUITADO
		'objEmail.TOs.Add destino,nombre_completo
		oMail.To = destino
		
		
		

		'especifica la dirección a la que se enviara una copia		
		if len(copia) > 0 then
		'Quitado 
		'objEmail.CCs.Add copia
		oMail.Cc = copia
		end if

		'especifica el tema del mensaje
		tema = objUpload.Form("tema")

		'Quitado 
		'objEmail.Subject = tema
		oMail.Subject =  tema

		'especifica el remitente del correo
		'Quitado 
		'objEmail.FromAddress = "aecys@ing.usac.edu.gt"

		oMail.From = "aecys@ing.usac.edu.gt"

		'specify an SMTP Relay server. This increases the speed and reliability of the operation
		'Quitado 
		'objEmail.SMTPRelayServers.Add "216.230.138.254"

		'initialize the HtmlBody property, we'll throw a header into it
		contenido = objUpload.Form("contenido")
		'Quitado 
		
		HTMLBody =  "<Html><Head></Head><Body><H2>" & contenido
		
		'oMail.Body = contenido
		'oMail.Body = HTML

		For Each Item in objUpload.Files
			'verifica que el campo del archivo sea el valido
			If (Item.TagName = "txtarchivo") Then
				
				'Quitado 		
				'	objEmail.HtmlEmbeddedObjs.Add Item.Path, 1, Item.OriginalPath
				HTMLBody = HTMLBody & "<a href=cid:1>Documento</a>"
				
				oMail.AttachFile(Item.Path)
					
			End If
			If (Item.TagName = "txtImage") Then
				If InStr(1,Item.ContentType,"image") Then
					'Quitado 		
					'objEmail.HtmlEmbeddedObjs.Add Item.Path, 2, Item.OriginalPath
					HTMLBody = HTMLBody & "<IMG SRC=cid:2>"
				oMail.AttachFile(Item.Path)

				End If
			End If
		Next

		'terminar html body colocando las etiquetas html finales

					'Quitado 				
			HTMLBody = HTMLBody & "</body></html>"

			oMail.Body = HTMLBody

			'oMail.AttachURL ("http://www.sat.gob.gt")
			'enviar el correo
			'QUITADO
			'objEmail.SendMail
			oMail.Send 

		'Prueba El exito/fallo
		If Err.Number <> 0 Then
		Response.Redirect ("previewforo.asp?desc="&"No se pudo enviar el mensaje, el siguiente error ocurrio: " & Err.Description)
		Else
		Response.Redirect ("previewforo.asp?desc="&opcion&"+The correo html fue enviado exitosamente. a "&destino)
		End If

		'Liberar recursos

		'QUITADO
		'Set objEmail = Nothing

		set oMail = Nothing


		Set objUpload = Nothing
		%>

