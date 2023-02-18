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
Set objEmail = Server.CreateObject("Dundas.Mailer") 

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


		
		'especifica el destinatario del mensaje
		objEmail.TOs.Add destino,nombre_completo
		
		

		'especifica la dirección a la que se enviara una copia		
		if len(copia) > 0 then
		objEmail.CCs.Add copia
		end if

'especifica el tema del mensaje
tema = objUpload.Form("tema")
objEmail.Subject = tema

'especifica el remitente del correo
objEmail.FromAddress = "aecys@ing.usac.edu.gt"

'specify an SMTP Relay server. This increases the speed and reliability of the operation
objEmail.SMTPRelayServers.Add "216.230.138.254"

'initialize the HtmlBody property, we'll throw a header into it
contenido = objUpload.Form("contenido")
objEmail.HTMLBody =  "<Html><Head></Head><Body><H2>" & contenido


For Each Item in objUpload.Files
	'verifica que el campo del archivo sea el valido
	If (Item.TagName = "txtarchivo") Then
		'If InStr(1,Item.ContentType,"audio") Then
			objEmail.HtmlEmbeddedObjs.Add Item.Path, 1, Item.OriginalPath
			objEmail.HtmlBody = objEmail.HtmlBody & "<a href=cid:1>Docu</a>"
		'End If
	End If
	If (Item.TagName = "txtImage") Then
		If InStr(1,Item.ContentType,"image") Then
			objEmail.HtmlEmbeddedObjs.Add Item.Path, 2, Item.OriginalPath
			objEmail.HtmlBody = objEmail.HtmlBody & "<IMG SRC=cid:2>"
		End If
	End If
Next

'terminar html body colocando las etiquetas html finales
objEmail.HTMLBody = objEmail.HTMLBody & "</body></html>"

'enviar el correo
'Response.Write "aqui<br>" 
'response.write "destino="&destino
'Response.Write "tema ="& tema
'Response.Write "contenido="&contenido
'Response.Write "htmlbody ="&objEmail.HTMLBody



'response.end 

objEmail.SendMail

'Prueba El exito/fallo
If Err.Number <> 0 Then
Response.Redirect ("previewforo.asp?desc="&"No se pudo enviar el mensaje, el siguiente error ocurrio: " & Err.Description)
Else
Response.Redirect ("previewforo.asp?desc="&opcion&"+The correo html fue enviado exitosamente. a "&destino)
End If

'Liberar recursos
Set objEmail = Nothing
Set objUpload = Nothing
%>

  
