<%@ Language=VBScript %>
<%
dim db
dim rs 
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select u.id,u.nombre,u.direccion,u.telefono,u.sitio,p.nombre pais,u.email from universidad u,pais p where u.id > 0 and u.pais = p.id order by u.pais, u.nombre"
set rs = db.SQLSelect("DBBAECYS",sql)

%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
function Regresar() {
 window.location  = "mant_universidades.asp"
 } 
 
 function RevisarSeleccion() {
	len = document.form.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.form.elements[i].name=='universidad') {
			if (document.form.elements[i].checked==true) 
			{
			   return true
			}
		}
	}
	alert("Error: No se selecciono ninguna universidad para actualizar.")
	return false
}

</script>
<BODY>

<div align="center">
<%
if not rs.eof then
%>
<form name="form" action="actualizar_universidad.asp" method="post" >
<table border="0" cellpadding="2" cellspacing="2">
      <tbody> 
      <tr class="celdaBgBlue"> 
        <th nowrap>&nbsp;</th>
        <th nowrap>Nombre</th>
        <th nowrap>Pais</th>
        <th nowrap>Direccion</th>
        <th nowrap>Telefono</th>
        <th nowrap>Sitio</th>
        <th nowrap>Correo Electronico</th>
      </tr>
<%
while rs.EOF <> true
id = rs("id")
nombre = rs("nombre")
direccion = rs("direccion")
telefono = rs("telefono")
sitio = rs("sitio")
pais = rs("pais")
email = rs("email")
%>
<tr class="celdaBgBlue05">
	  <td class="celdaBgBlueLeft">
		<input type="radio" name="universidad" value="<%=id%>">
        </td>
	  <td><% if len(Nombre) > 0 then Response.Write Nombre else Response.Write (" Dato no ingresado") end if%></td>
	  <td><% if len(pais) > 0 then Response.Write Pais else Response.Write (" Dato no ingresado") end if%></td>
	  <td><% if len(direccion) > 0 then Response.Write Direccion else Response.Write (" Dato no ingresado") end if%></td>
	  <td><% if len(Telefono) > 0 then Response.Write telefono else Response.Write (" Dato no ingresado") end if%></td>
	  <td><% if len(sitio) > 0 then Response.Write Sitio else Response.Write (" Dato no ingresado") end if%></td>
	  <td><% if len(emial) > 0 then Response.Write email else Response.Write (" Dato no ingresado") end if%></td>
</tr>
<%
rs.MoveNext 
wend%>
  </tbody>
</table>
<br>
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="Actualizar" type ="submit" value="Actualizar" onClick="return RevisarSeleccion()"></td>
<td width="17"></td>
<td><input type="button" name="Cancelar" value="Cancelar" onclick="javascript:Regresar()"></td>
</tr>
</table>
</form>
<%
else
%>
<h2>No hay Universidades Ingresadas</h2>
<br>
<a href="mant_universidades.asp">Regresar...</a>
<%
end if
set rs = nothing

%>

</div>
</BODY>
</HTML>

