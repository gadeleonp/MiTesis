<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim rs 
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select p.id,p.nombre from pais p order by p.nombre"
set rs = db.SQLSelect ("DBBAECYS",sql)

%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">

 function RevisarSeleccion() {
	len = document.form.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.form.elements[i].name=='pais') {
			if (document.form.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("Error: No se selecciono ningun pais para actualizar.")
	return false
}
function Regresar() {
window.location  = "mant_paises.asp"
  }
</script>
<BODY>

<div align="center">
<%
if not rs.eof then
%>
<form name="form" action="actualizar_pais.asp" method="post" onsubmit="return RevisarSeleccion()">
<table border="0" cellpadding="2" cellspacing="2">
      <tbody> 
      <tr> 
        <td class="celdaBgBlue">-</td>
        <td class="celdaBgBlue">Nombre</td>
      </tr>
<%
while rs.EOF <> true
id = rs("id")
nombre = rs("nombre")
%>
<tr>
	  <td class="celdaBgBlueLeft">
		<input type="radio" name="pais" value="<%=id%>">        </td>
	  <td class="celdaBgBlue05"><% =Nombre%></td>
</tr>
<%
rs.MoveNext 
wend%>
  </tbody>
</table>
<br>
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="Actualizar" type ="submit" class="celdaBgBlue05" value="Actualizar"></td>
<td width="17"></td>
<td><input name="Cancelar" type="button" class="celdaBgBlue05" onclick="javascript:Regresar()" value="Cancelar"></td>
</tr>
</table>
</form>
<%
else
%>
<h2>No hay Paises Ingresados</h2>
<br>
<a href="mant_paises.asp">Regresar...</a>
<%
end if
set rs = nothing
set db = nothing
%>

</div>
</BODY>
</HTML>

