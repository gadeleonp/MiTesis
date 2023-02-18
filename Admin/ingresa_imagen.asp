<%@ Language=VBScript %>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
function cerrarventana() {
	   window.close()
 	}

</script>
<BODY>
<%
dim noticia
dim folder
dim objUpload

dim rs

set objUpload = server.CreateObject ("Dundas.upload.2")
objUpload.MaxFileCount = 6

objUpload.UseUniqueNames = false
objUpload.UseVirtualDir = true

set objNextFile = objUpload.GetNextFile 

folder = objUpload.Form("folder")
noticia = objUpload.Form("noticia")
folder_imagen = objUpload.Form("folder_img")


TempFolderName = "/Noticias/" & folder & "/" & folder_imagen
'Response.Write tempfoldername
'Response.end
objUpload.DirectoryCreate TempFolderName,false


response.write "Folder = "&tempFolderName


Do Until objNextFile Is Nothing

If InStr(1,objNextFile.ContentType,"octet-stream",1) <> 0 Then
	If InStr(1,objNextFile.ContentType,"image",1) <> 0 Then
%>
<h4>El archivo "<%=objNextFile.FileName %>" no fue grabado por ser una aplicacion.</h4>
  <%
	end if
End If

If InStr(1,objNextFile.ContentType,"image",1) <> 0 Then
objnextfile.save TempFolderName

%>
<h4>El archivo "<%= objNextFile.FileName %>" fue almacenado con exito.</h4>
<h4>Tipo "<%= objNextFile.ContentType %>"</h4>
<%
else
%>

<h4>El archivo "<%=objNextFile.FileName %>" no fue grabado por no ser una imagen.</h4>

<%
End If

Set objNextFile = objUpload.GetNextFile 
Loop 
set objUpload = nothing

dim db
dim resultado
dim sql
 set db = server.CreateObject ("SQLComandos.Comandos")
 sql = "update noticia set folder_imagenes = '"&folder_imagen&"' where id = "&noticia
 resultado = db.SQLOtras("DBBAECYS",sql) 
%>
<form name="imagen" method = "post" action="agregar_imagen.asp">
<input name="noticia" type="hidden" value ="<%=noticia%>">
<input name="folder" type="hidden" value ="<%=folder%>">
<table border="0" align="center">
<tr>
<td><input name="Cerrar" type="submit" value="Cerrar" onclick="return cerrarventana()"></td>
<td><input name="Mas Imagenes" type="submit" value="(+) Imagenes"></td>
</tr>
</table>
</form>
</BODY>
</HTML>
