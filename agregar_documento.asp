<%@ Language=VBScript %>
<%
if isempty(session("id")) then
%>
<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
</head>
	<script language="javascript">
	 function cerrarventana() {
	   opener.window.location = 'Login.asp'
	   window.close
	 }
	</script>
<body onload="javascript:cerrarventana()">
</body>
</html>
<%
end if

Dim id
dim login
dim tipo

id = session("id")
respuesta = request.QueryString("respuesta")
foro = Request.QueryString("foro")
enlace = Request.QueryString("enlace")
tipo = ""
tipo = request.querystring("tipo")
dim db
set db = server.CreateObject("SQLComandos.Comandos")
%>

<HTML>
<HEAD>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<title>
Ingreso Documento
</title>
</HEAD>
<script language="javascript">
function cerrarventana() {
	   window.close()
 	}
 	
function Validar(Forma) {
if ((Forma.documento.value.length == 0) && (Forma.action == "ingresar_documentoforo.asp"))
{
	alert ("No ingreso archivo")
	return false
}
return true
}
</script>
<BODY>
<%
dim sql
dim rs


%>
<%
if  tipo = "2" and not isempty(respuesta) then 
sql = "select count(id) contador from documento where respuesta = " & respuesta

else
sql = "select count(id) contador from documento where foro = " & foro
end if


set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
if rs("contador") < 3 then
%>
<div align="center">
<form  name="ingresar_documento" method="post" enctype="multipart/form-data" action="ingresar_documentoforo.asp" >
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td class="celdaBgBlueLeft">Documento: </td>
<td class="celdaBgBlue05">
<input name="documento" type="file">
</td>
</tr>
</table>
<br>
<%
if not isempty(tipo) and tipo="2" then
%>
<input name="respuesta" type="hidden" value="<%=respuesta%>">
<input name="tipo" type="hidden" value="<%=tipo%>">
<input name="foro" type="hidden" value="<%=foro%>">
<%
else
%>
<input name="foro" type="hidden" value="<%=foro%>">
<%
end if
%>
<input name="enlace" type="hidden" value="<%= enlace%>">
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="aceptar" type="submit" class="celdaBgBlue05" onClick="return Validar(ingresar_documento)" value="Aceptar"> </td>
<td><input name="cancelar" type="submit" class="celdaBgBlue05" onClick="return cerrarventana()" value="Cancelar"></td>
</tr>
</table>
</form>
</div>
<%
else
%>
<div align="center">
<h2 >No se pueden Ingresar mas de 3 documentos</h2>
<input name="cancelar" type="submit" class="celdaBgBlue05" onClick="return cerrarventana()" value="Cancelar">
</div>
<%
end if
end if
set rs = nothing
%>
<%
if  tipo = "2" and not isempty(respuesta) then 
sql = "select enlace,archivo from documento where respuesta = "&respuesta
else
sql = "select enlace,archivo from documento where foro = "&foro
end if
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
<br>
<div align="center">
<h2>Documentos Ingresados</h2>
<table border="0" cellpadding="2" cellspacing="2">
<tbody>
<tr>
<th align="center" class="celdaBgBlue">&nbsp;Enlace&nbsp;</th>
<th align="center"  class="celdaBgBlue" >Archivo</th>
</tr>
<%
while rs.eof <> true 
%>
<tr align="center">
<td class="celdaBgBlue05"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=rs("enlace")%>&nbsp;</font></td>
<td  class="celdaBgBlue05"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=rs("archivo")%>&nbsp;</font></td>
</tr>
<%
rs.movenext 
wend
%>
</tbody>
</table>
</div>
<%
end if
set rs = nothing
%>
</BODY>
</HTML>
<%
set db = nothing
%>