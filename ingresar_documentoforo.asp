<%@ Language=VBScript %>
<%
if isempty(session("id")) then
    Response.Redirect("Login.asp")
end if
dim db 
dim resultado
set db = server.CreateObject("SQLComandos.Comandos")
%>

<HTML>
<HEAD>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<script>

function Ingresar() {
  if (documento.tipo.value == '2' )  {
//  alert(documento.tipo.value+' si entro')
	  documento.submit()
  }
  else
  {
    opener.document.location = "reporte_foros.asp"
  }
	if (window && !window.closed) 
	{
		   window.close()
    }
  }
</script>

</HEAD>
<BODY onload="return Ingresar()" >
<div align="center">
<%
dim objupload
set objupload = server.CreateObject ("dundas.upload.2")


Set objNextFile = objUpload.GetNextFile
objupload.UseVirtualDir = true


If Err.Number <> 0 Then Response.Redirect "Error.asp"
	
	

	Path =  "/aecys/paecycod/Foros"
	objupload.DirectoryCreate Path
	objupload.UseUniqueNames = true


Do Until objNextFile Is Nothing
	If InStr(1,objNextFile.ContentType,"octet-stream",1) = 0 Then
		objnextfile.save Path
	else
	%>
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
tipo = objupload.form("tipo")
respuesta = objupload.form("respuesta")
foro = objupload.Form("foro")
enlace = objupload.Form("enlace")



set objupload = nothing



if tipo = "2" and not isempty(respuesta) then
sql = "insert into documento (foro,respuesta,archivo,enlace) values (" & foro & ","& respuesta & ",'" & arc_curriculum &"','" & enlace & "')"
else
sql = "insert into documento (foro,archivo,enlace) values (" & foro & ",'" & arc_curriculum &"','" & enlace & "')"
end if
resultado = db.SQLOtras("DBBAECYS",sql)
set db = nothing
%>
<%
if tipo= "2" then
%>
<form name="documento" action="desarrollo_foro.asp" target="principal" method ="get">
	<input name="tipo" type="hidden" value="<%=tipo%>">	
	<input name="foro" type="hidden" value="<%=foro%>">	
</form>
<%
else
%>
<form name="documento" action="nuevoforo.asp" method ="get">
	<input name="tipo" type="hidden" value="<%=tipo%>">	
	<input name="archivo" type="hidden" value="<%=arc_curriculum%>">
	<input name="descripcion" type="hidden" value="<%=descripcion%>">
	<input name="enlace" type="hidden" value="<%=enlace%>">	
</form>
<%
end if
%>
</div>
</BODY>
</HTML>