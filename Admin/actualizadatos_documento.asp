<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db
dim objupload
dim objNextFile
dim resultado
dim rs
set objupload = server.CreateObject ("dundas.upload.2")

set db  = server.CreateObject ("SQLComandos.Comandos")

objupload.UseUniqueNames = false
objupload.UseVirtualDir = true


set objNextFile = objupload.GetNextFile 


'Datos de la noticia
	noticia = objupload.Form("noticia")
	documento = objupload.Form("documento")
	enlace_anterior = objupload.Form("enlace")
	
	
	sql = "select folder_contenedor from noticia where id ="& noticia
	set rs = db.SQLSelect("DBBAECYS",sql)
	if not rs.EOF  then
		folder = rs("folder_contenedor")
	end if
	set rs = nothing


TempFolderName =  "/aecys/paecycod/Noticias/" & folder
'Response.Write folder & " " & noticia & " " & documento & " " & TempFolderName & "\"

objUpload.DirectoryCreate TempFolderName,false
 
Do Until ObjNextFile is nothing

If InStr(1,objNextFile.ContentType,"octet-stream",1) = 0 Then
 objnextfile.save TempFolderName
End If
Set objNextFile = objUpload.GetNextFile 
Loop




for each item in objUpload.Files

if len(item.TagName) = 7 then
arc_noticia = objupload.GetFileName (item.Path)
ctype_noticia  =Item.ContentType 
Response.Write "arc_noticia = " & arc_noticia & " tipo = "& ctype_noticia &"<br>"
band1 = 1
ingresado1 = 1
end if
if len(item.TagName) = 10 then
arc_img1 = objupload.GetFileName (item.Path)
ctype_img1= Item.ContentType 
Response.Write "arc_imagen1 = " & arc_img1 & " tipo = "& ctype_img1 &"<br>"
band2 = 1
ingresado2 = 1
end if

if len(item.TagName) = 11 then
arc_img2 = objupload.GetFileName (item.Path)
ctype_img2= Item.ContentType 
Response.Write "arc_imagen2 = " & arc_img2 & " tipo = "& ctype_img2 &"<br>"
band3 = 1
ingresado3 = 1
end if
Next

if band1 = 1 then
		
		If InStr(1,ctype_noticia,"text/htm",1) <> 0 Then
			
		else
		    objupload.Files.Item(0).Delete()
		    Response.Write "Archivo Eliminado por no ser html " & arc_noticia & "<br>"
		    ingresado1 = 0
		end if
	end if
	
	if band2 = 1 then
		If InStr(1,ctype_img1,"image",1) <> 0 Then
		else
		    if band1 = 1 then
		    objupload.Files.Item(1).Delete()
		    ingresado2 = 0
		    else
		    objupload.Files.Item(0).Delete()
		    ingresado2 = 0
		    end if
		    Response.Write "Archivo Eliminado por no ser Imagen " & arc_img1 & "<br>"
		end if
	end if
	
	if band3 = 1 then
		If InStr(1,ctype_img2,"image",1) <> 0 Then
		else
			if band1 = 0 and band2 = 0 then
				objupload.Files.Item(0).Delete()								
				ingresado3 = 0
			end if	
			if (band1 = 0 and band2 = 1) or (band1 = 1 and band2 = 0) then
				objupload.Files.Item(1).Delete()			
				ingresado3 = 0
			end if
			if (band1 = 1 and band2 = 2) then
				objupload.Files.Item(2).Delete()
				ingresado3 = 0
			end if
		    Response.Write "Archivo Eliminado por no ser Imagen " & arc_img2 & "<br>"
		end if
	end if

sql = "select enlace from documento_noticia where noticia = "& noticia & " and id = " &documento
set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.EOF  then
enlace = rs("enlace")
objupload.FileDelete (TempFolderName+"/"+enlace)
end if
set rs = nothing
sql = "update documento_noticia "

if ingresado1 = 1 then
Sql = Sql & " set enlace = '" & arc_noticia & "'"
end if

if ingresado2 = 1 then
Sql = Sql & ", imagen_principal = '" & arc_img1 & "'"
end if


if ingresado3 = 1 then
Sql = Sql & ", imagen_secundaria = '" & arc_img2 & "'"
end if

Sql = Sql & " where noticia = " & noticia  & " and id = " & documento
Response.Write "<br><br>"
Response.Write sql
resultado = db.SQLOtras("DBBAECYS",sql)
set db = nothing
'release resources
Set objUpload = Nothing
%>
<html>
<script language="javascript">
 function Continuar(Forma) {
 Forma.submit()
 return true
 }
</script>
<body onload ="return Continuar(redireccion)">
<form name="redireccion" action="actualizar_noticia.asp" method = "post">
  <input name="noticia" type ="hidden" value ="<%=noticia%>">
</form>
</body>
</html>