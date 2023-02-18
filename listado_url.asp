<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
<link href="styles/interior.css" rel="stylesheet" type="text/css">
<script language="javascript">
function Enlace(url)
{
 window.location = url
}


</script>
</head>
<body>
<%
	 if not rs.eof then
%>
<div align="left">
<table align="left" border="0" cellpadding="2" cellspacing="2">
<%
		while rs.EOF <> true
%>

<tr>
<td align="left"><a title='<%=rs("descripcion")%>' target="_top" href="<%=rs("url")%>"><font size="-1"><%=rs("url")%></font></a></td>
</tr>
<%		rs.movenext
		wend	
%>
</table>
</div>
<%
else
%>
<h4>No se han ingresado enlaces de interes.</h4>
<%
end if
%>
<br>
</body>
</html>
<%
set rs = nothing
set db = nothing
%>
