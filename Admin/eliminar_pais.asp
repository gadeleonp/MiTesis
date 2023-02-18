<%@ Language=VBScript %>

<%
dim db
dim rs 
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select p.id, p.nombre from pais p order by p.nombre"
set rs = db.SQLSelect("DBBAECYS",sql)

%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
function Regresar() {
 window.location  = "mant_paises.asp"
  }
  
function eliminar(){
	if (confirm("¿Desea eliminar los elemenos seleccionados?\n Los cambios serán irreversibles."))
	{
	 return true;
	}
	else
	{
		return false;
	}
}
  
   function RevisarSeleccion() {
	len = document.form.elements.length;
	var i=0;
	for( i=0; i<len; i++) {
		if (document.form.elements[i].name=='seleccion') {
			if (document.form.elements[i].checked==true) 
			{
			   eliminar()
			   return true
			}
		}
	}
	alert("Error: No se selecciono ningun pais para eliminar.")
	return false
}

</script>
<BODY>

<div align="center">
<%
if not rs.EOF then
%>
<form name="form" action="elimina_pais.asp" method="post" onsubmit="return RevisarSeleccion()">
<table border="0" cellpadding="2" cellspacing="2">
  <tbody>
    <tr class="celdaBgBlue">
    <th nowrap>-</th>
	  <th nowrap>Nombre</th>
	 </tr>
<%
while rs.EOF <> true
id = rs("id")
nombre = rs("nombre")
%>
<tr>
	  <td class="celdaBgBlueLeft"><input name="seleccion" type="checkbox" value="<%=id%>"></td>
	  <td class="celdaBgBlue05"><% =Nombre %></td>
</tr>
<%
rs.MoveNext 
wend%>
  </tbody>
</table>
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="Eliminar" type ="submit" value="Eliminar" class="celdaBgBlue05"></td>
<td width="17"></td>
<td><input type="button" name="Cancelar" value="Cancelar" onclick="return Regresar()" class="celdaBgBlue05"></td>
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

