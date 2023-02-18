<%@ Language=VBScript %>
<%

if isempty(session("id")) then
    Response.Redirect("Login.asp")
end if
dim db
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
dim resultado
dim objupload
set objupload = server.CreateObject ("dundas.upload.2")

Set objNextFile = objUpload.GetNextFile

objupload.UseVirtualDir = true

If Err.Number <> 0 Then Response.Redirect "Error.asp"
	
	

        Path = "/aecys/paecycod/Usuarios"
	objupload.DirectoryCreate Path
	objupload.UseUniqueNames = true
        

Do Until objNextFile Is Nothing
	If InStr(1,objNextFile.ContentType,"octet-stream",1) = 0 Then
		objnextfile.save Path 
	else
	%>
	<html>
	<head>
	<link href="./theme/Master.css" rel="stylesheet" type="text/css">
	</head>
	<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >
	<hr> 
	<h2>Error: El archivo que intento subir es una aplicacion.<br> Operacion Cancelada.</h2>
	<hr>
	</div>
</BODY>
</HTML>
	<%
	End If
	Set objNextFile = objUpload.GetNextFile 
Loop 

for each item in objUpload.Files

arc_curriculum = objupload.GetFileName (item.Path)
ctype_curriculum  =Item.ContentType 


next
login = objupload.Form("login")
folder = objUpload.Form("folder")

sql = " select arc_resume from usuario where id = '"& session("id") &"'"

set rs = db.SQLSelect ("DBBAECYS", sql)

path = path +"/" +rs("arc_resume")

	if rs.eof <> true then
		if len(path) > 0 and len (rs("arc_resume")) > 0then
			objupload.filedelete path
		end if
'		Response.Write "entro"

	else
'	Response.Write "no entro"
	end	if
	
	
	set rs = nothing
sql = "update usuario set arc_resume = '" & arc_curriculum & "' where  id = "& session("id")

resultado = db.SQLOtras("DBBAECYS",sql)
%>
<%



dim menu
menu = objupload.Form("menu")
%>
<HTML>
<HEAD>
</HEAD>
<%
if menu <> "1" then
%>
<script>
function Ingresar() {
	valor = "ingreso_preferencias.asp?archivo="+resume.archivo.value
	opener.document.location = valor
	if (window && !window.closed) 
	{
		   window.close()
	 }
}
</script>
<%
else
%>
<script>
function Ingresar() {
	valor = "ingreso_preferencias.asp?archivo="+resume.archivo.value
	window.location = valor
}
</script>	
<%
end if
set objupload = nothing
%>	 

<BODY onload="return Ingresar()" >
<div align="center">


<form name="resume" action="ingreso_preferencias.asp" method ="post">
<input name="archivo" type="hidden" value="<%=arc_curriculum%>">
</form>
</div>
</BODY>
</HTML>