<%@ Language=VBScript %>
<%
dim usuario
usuario = session("id")
if isempty(usuario) then
%>
<html>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
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
foro = Request.QueryString("foro")
tipo = request.QueryString("tipo")
respuesta = request.QueryString("respuesta")
dim db
dim rs
set db = server.CreateObject("SQLComandos.Comandos")

if tipo="2" and not isempty(respuesta) then
sql = "select id,enlace,archivo from documento where respuesta = "&respuesta & "and foro = "& foro
else
sql = "select id,enlace,archivo from documento where foro = "&foro
end if
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>

<html>
<head>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Eliminar Documento</title>
</head>
<script language="JavaScript">
	function cerrarventana() {
	window.close()
  	//	   window.parent.close()
  		   
	  	   
	 }
 function Eliminar(Documento) {
    forma.documento.value=Documento
   forma.submit()
 }
</script>
<body>
<form name="forma" method="post" action="elimina_docforo.asp">
<input name="documento" type="hidden" value="">
<input name="tipo" type="hidden" value="<%=tipo%>">
</form>
<br>
<div align="center">
<h2>Documentos Ingresados</h2>
<table border="0" cellpadding="2" cellspacing="2">
<tbody>
<tr>
<th  class="celdaBgBlue" align="center">&nbsp;Enlace&nbsp;</th>
<th  class="celdaBgBlue" align="center">Archivo</th>
<th  class="celdaBgBlue" align="center">&nbsp;</th>
</tr>
<%
while rs.eof <> true 
%>
<tr  class="celdaBgBlue05" align="center">
	<td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=rs("enlace")%>&nbsp;</font></td>
	<td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=rs("archivo")%>&nbsp;</font></td>
	<td>&nbsp;<img alt="Eliminar" src="iconos/deletes.gif" onclick="javascript:Eliminar('<%=rs("id")%>')" WIDTH="22" HEIGHT="19">&nbsp;</td>
</tr>
<%
rs.movenext 
wend
%>
</tbody>
</table>
<table width="0%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><input name="Cancelar"  class="celdaBgBlue05" type="button" onClick="javascript:cerrarventana()" value="Cancelar"></td>
  </tr>
</table>

</div>
</body>
</html>
<%
else
%>
<html>
<script language="javascript">
	function cerrarventana() 
	{
		alert("Este foro no tiene documentos enlazados.\n")
		window.close()	   
 	}
</script>
<body onload="javascript:cerrarventana()">
</body>
</html>
<%
end if
set rs = nothing
set db = nothing
%>