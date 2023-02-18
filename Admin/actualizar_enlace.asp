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
 sql = "select id,url,texto, descripcion from enlace"
 
 set rs = db.sqlselect("DBBAECYS",sql)

%>
<html>
<head>
<title>Actualizar Enlace</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</head>
<script language="javascript">
function Eliminar() {
if (RevisarSeleccion())
 {
 document.form1.submit()
 }
}
function CheckAll(checked) {
	len = document.form1.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.form1.elements[i].name=='id') {
			document.form1.elements[i].checked=checked;
		}
	}
}
function RevisarSeleccion() {
	len = document.form1.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.form1.elements[i].name=='id') {
			if (document.form1.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("No se selecciono ninguna noticia para actualizar.")
	return false
}
function Menu()
{
 document.location = "mantenlaces.asp"
}
</script>
<body>
<h1 align="center" >Actualizar Enlace</h1>
<br>
<form name="form1" method="post" action="actualiza_enlace.asp">
<br>
<%
	 if not rs.eof then
%>
<table align="center" border="0" cellpadding="2" cellspacing="2">
<tr class="celdaBgBlue">
<th>--</th><th>Enlace</th><th>Texto</th>
</tr>
<%
		while rs.EOF <> true
%>
<tr>
<td class="celdaBgBlue05"><input name="id" type="radio" value="<%=rs("id")%>"></td>
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
<td><input name="Aceptar" type="button" value="Aceptar" class="celdaBgBlue05" onclick="javascript:Eliminar()"></td>
<td width="17"></td>
<td><input name="Cancelar" type="button" value="Cancelar" class="celdaBgBlue05" onclick="javascript:Menu()"></td>
</tr>
</table>
    </form>
</body>
</html>
<%
set rs = nothing
set db = nothing
%>