<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%

'Datos de la noticia
dim titulo
dim resumen
dim fecha
dim autor
dim arc_notica
dim arc_img1
dim arc_img2
'Datos de la noticia

dim objUpload		
dim NextFile
dim strMessage		
dim strPath		

set objUpload = Server.CreateObject("Dundas.Upload")

objUpload.MaxFileCount = 3
'objUpload.MaxFileSize = objUpload.MaxFileSize * 1024
'objUpload.DirectoryCreate TempFolderName
objUpload.UseUniqueNames = false

nombrefolder = "Transicion"

RootFolderName = Server.MapPath("..") & "\"
TempFolderName = RootFolderName & "Noticias\" & nombrefolder



objUpload.DirectoryCreate TempFolderName






set NextFile = objUpload.getNextFile

cont = 1

do until NextFile is nothing

	NextFile.Save TempFolderName

	'get Nextfile object
	set NextFile = nothing
	
	set NextFile = objUpload.GetNextFile
	Nextfile.save = TempFolderName
	
	cont = cont + 1 
loop


Response.Write "Numero de archivos subidos: " & CStr(objUpload.Files.Count) 
it = 0

valor = objUpload.Files.Item.GetAttributes("filename")
Response.Write "valor = " & valor
for each item in objUpload.Files
	
		
'	Response.Write "<bR>Input Tag name:"
'	Response.Write item.TagName & "<bR>"
'	Response.Write "Content type: "
'	Response.Write item.ContentType & "<bR>"
'	Response.Write "Size: "
'	Response.Write CStr(item.Size) & "<bR>"
'	Response.Write " Original path: "
'	Response.Write item.OriginalPath & "<bR>"
'	Response.Write "Server path: "
'	Response.Write item.Path & "<bR>"
'	Response.Write "Archivo: "
'	Response.Write objUpload.GetFileName (item.Path) & "<bR>"
'	Response.Write "Directorio: "
'	Response.Write objUpload.GetFileDirName  (item.Path) & "<bR>"
	
	
	
	
	'item.Delete		' Delete uploaded file
	
	it = it + 1

	arc_notica = ""
	arc_img1 = ""
	arc_img2 = ""
	
	if it = 1 then
    arc_noticia	= objUpload.GetFileName (item.Path)
	end if
	
	if it = 2 then
	arc_img1 = objUpload.GetFileName (item.Path)
	end if
	
	if it = 3 then
	arc_img2 = objUpload.GetFileName (item.Path)
	end if
	
next

'show form items collection

it = 0

'delete temp folder
'objUpload.DirectoryDelete TempFolderName, true
call CheckErr

'destroy Upload component instance

titulo = objUpload.Form ("txttitulo")
descripcion = objUpload.Form("txtdescripcion")
autor = objUpload.Form("txtautor")



Response.Write "titulo=" & titulo & "<br>"
Response.Write "autor=" & autor & "<br>"
Response.Write "descripcion=" & descripcion & "<br>"
if len(arc_noticia) = 0 then
Response.Write("No se Ingreso la Noticia")
else
	Sql = "insert into noticia (titulo,resumen,enlace,imagen_principal,imagen_secundaria,fecha_publicacion,autor,folder_contenedor) values ('" & titulo & "','" & descripcion & "','" & arc_noticia & "','" & arc_img1 & "','" & arc_img2 & "','" & now  &  "','" & autor & "','"& nombrefolder &"')"

 end if
 Response.Write(sql)
 
set objUpload = nothing

%>
<html>
<body>
</body>
</html>