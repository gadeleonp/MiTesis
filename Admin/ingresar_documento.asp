<%@ Language=VBScript %>
<%
dim db
dim resultado
 set db = server.CreateObject ("SQLComandos.Comandos")
 

dim arc_notica
dim arc_img1
dim arc_img2
'Datos de la noticia

dim objUpload		
dim strMessage		
dim strPath		
'Creacion del objeto manejador de archivos
set objUpload = Server.CreateObject("Dundas.Upload")



objUpload.MaxFileCount = 3
objUpload.UseUniqueNames = false

'objUpload.MaxFileSize = objUpload.MaxFileSize * 1024
'objUpload.DirectoryCreate TempFolderName



set objNextFile = objUpload.getNextFile
'Creacion del Folder Contenedor de la noticia
objUpload.UseVirtualDir = true
noticia = objUpload.Form ("noticia")
folder = objUpload.Form("folder")


TempFolderName = "/aecys/paecycod/Noticias/" & folder


'Creacion de Directorio
objUpload.DirectoryCreate TempFolderName,false






cont = 1

Do Until objNextFile Is Nothing

If InStr(1,objNextFile.ContentType,"octet-stream",1) = 0 Then
	objnextfile.save TempFolderName
	Response.write "Tipo " & objnextfile.ContentType & "<br>"
	Response.Write objNextFile.TagName & "<br>"
	cont = cont + 1 
else
    Response.Write "El archivo : '" & objNextFile.FileName & "' No fue grabado por ser una aplicacion."
End If

	
	Set objNextFile = objUpload.GetNextFile 

Loop 	


Response.Write "Numero de archivos subidos: " & CStr(objUpload.Files.Count) 

	it = 0
	arc_notica = ""
	arc_img1 = ""
	arc_img2 = ""
	


for each item in objUpload.Files



	it = it + 1

	
	if len(item.tagname) = 7 then
    arc_noticia	= objUpload.GetFileName (item.Path)
    Response.Write item.tagname
    Response.Write "arc_noticia = " & arc_noticia & "<br>"
    
	end if
	
	if len(item.tagname) = 10  then
	Response.Write item.tagname
	arc_img1 = objUpload.GetFileName (item.Path)
	    Response.Write "arc_img1 = " & arc_img1 & "<br>"
	end if
	
	if len(item.tagname) = 11 then
	Response.Write item.tagname
	arc_img2 = objUpload.GetFileName (item.Path)
	Response.Write "arc_img2 = " & arc_img2 & "<br>"
	end if
	
next

'show form items collection

it = 0

'delete temp folder
'objUpload.DirectoryDelete TempFolderName, true

'destroy Upload component instance

dim rs 


sql = "select max(correlativo)+1 correlativo from documento_noticia where noticia = " & noticia
Response.Write sql
set rs = db.SQLSelect("DBBAECYS",sql)

if rs.eof = true then
else
	correlativo = rs("correlativo")
end if

	if len(arc_noticia) = 0 then
		Response.Write("No se Ingreso la Noticia")
	else
		Sql = "insert into documento_noticia (noticia,correlativo,enlace,imagen_principal,imagen_secundaria) values (" & noticia & "," & correlativo &",'" & arc_noticia & "','" & arc_img1 & "','" & arc_img2 & "')"
		Response.Write "<br>noticia = " & noticia
		Response.Write sql &"-: "&correlativo
		
		resultado = db.SQLOtras("DBBAECYS",sql)
	
	end if

	
	//Response.Write(sql)
	 
set objUpload = nothing
Response.Redirect("confirma_documento.asp?noticia="& noticia & "&enlace="&arc_noticia)
%>
