<%
		Dim objUpload 'stores an Upload control object
		Dim objEmail 'stores a Mailer control object
		Dim strPath 'stores path to a directory we create at same level as this ASP page 
		dim autor
		dim opcion
		dim db

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


'		opcion = objUpload.Form("opcion")

		autor = objUpload.Form("autor")
		copia = objUpload.Form("copia")
		tema = objUpload.Form("tema")
		contenido = objUpload.Form("contenido")
		destino = objUpload.Form("lstus")



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
			
			''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''
		sql = "select mail from usuario where id in (" & destino & ") and recibir_mail = 'SI'"
		set rs = db.SQLSelect("DBBAECYS",sql)
		correo  = ""
	if not rs.eof then
		while rs.eof <> true
			correo = correo & rs("mail")  & ","
		

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
	
	
	correo = left(correo,(len(correo)-1))
	
	oMail.To = correo
			
			''''''''''''''''''''''''''''''''''
			'''''''''''''''''''''''''''''''''''
			oMail.Send 


		set oMail = Nothing
		Set objUpload = Nothing
		set rs = nothing
		set	db = nothing


		'Prueba El exito/fallo
		If Err.Number <> 0 Then
		Response.Redirect ("./previewforo.asp?desc="&"No se pudo enviar el mensaje, el siguiente error ocurrio: " & Err.Description)
		Else
		Response.Redirect ("./previewforo.asp?desc="&opcion&"El correo html fue enviado exitosamente.")
		End If


		%>

