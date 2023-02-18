<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db
dim rs

set db = server.CreateObject ("SQLComandos.Comandos")
dim universidad

pais = Request.Form("pais")

%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<title>Datos del Pais</title>
</HEAD>
<script language="javascript">
function Regresar() {
 form1.action = "listado_pais.asp"
 form1.submit
 return true
 }
</script>
<BODY>
<%
sql = "select id,nombre  from pais where id = " & pais
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.EOF then
id = rs("id")
nombre = rs("nombre")
end if
set rs = nothing
%>

<h2 align="center">Actualizar datos pais</h2>
 <form name="form1" method="post" action="actualiza_pais.asp">
<table align="center" width="75%" border="0" cellpadding="2" cellspacing="2">
  <tr>
    <td width="29%" class="celdaBgBlueLeft">Nombre del Pais</td>
    <td width="71%" class="celdaBgBlue05">
        <input type="text" name="nombre" value="<%=nombre%>">
   </td>
  </tr>
</table>
<br>
<table align="center" border="0" cellpadding="2" cellspacing="2">
  <tbody>
  <tr>
    <td><input type="submit" name="Aceptar" value="Aceptar" class="celdaBgBlue05"></td>
	<td width="17"></td>
    <td><input type="hidden" name="id" value="<%=id%>"></td>
	    <td><input type="submit" name="Cancelar" value="Cancelar" class="celdaBgBlue05" onclick="return Regresar()"></td>
  </tr>
  </tbody>
</table>
      </form>
</BODY>
</HTML>
<%
set db = nothing
%>