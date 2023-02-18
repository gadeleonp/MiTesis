<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim rs
dim sql
 set db = server.CreateObject("sqlcomandos.comandos")
 sql = "select url,texto, descripcion from enlace"
 
 set rs = db.sqlselect("DBBAECYS",sql)
%>
<html>
<head>
<title>Enlaces de Interes</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
function Enlace(url)
{
 document.location = url
}
function Menu()
{
 document.location = "mantenlaces.asp"
}

</script>
</head>
<body>
<%
	 if not rs.eof then
%>
<h2 align="center">Enlaces de Interes Ingresados</h2>
<br>
<table align="center" border="0" cellpadding="2" cellspacing="2">
<tr class="celdaBgBlue">
<th>Sitio</th><th>Descripcion</th>
</tr>
<%
		while rs.EOF <> true
%>
<tr>
<td class="celdaBgBlue05"><a  title='<%=rs("descripcion")%>' href="javascript:Enlace('<%=rs("url")%>')"><%=rs("url")%></a></td>
<td class="celdaBgBlue05"><span title="<%=rs("descripcion")%>" ><%=rs("texto")%></span></td>
</tr>
<%		rs.movenext
		wend	
%>
</table>
<%
else
%>
<h2>No se han ingresado enlaces de interes.</h2>
<%
end if
%>
<br>
<table align="center" border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="Aceptar" type="button" value="Aceptar" class="celdaBgBlue05" onclick="javascript:Menu()"></td>
</tr>
</table>
</body>
</html>
<%
set rs = nothing
set db = nothing
%>
