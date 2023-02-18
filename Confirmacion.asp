<%@ Language=VBScript %>
<%
Ingresado = Request.Item ("Ingresado")

if Ingresado = 1 then
  descripcion = "Los Datos de la Universidad han sido ingresados correctamente."
  destino = "mant_universidades.asp"
end if

if Ingresado = 2 then
  descripcion = "Los Datos del pais han sido ingresados correctamente"
  destino = "mant_paises.asp"
end if

if Ingresado = 3 then
  descripcion = "La categoria fue ingresada correctamente"
  destino = "mant_areas.asp"
  end if
  
  if Ingresado = 4 then
  descripcion = "El aera fue ingresada correctamente"
  destino = "mant_areas.asp"
  end if
  
  if Ingresado = 5 then
  descripcion = "La pregunta de la encuesta fue ingresada correctamente"
  destino = "mantencuesta.asp"
  end if
%>
<HTML>
<HEAD>
<H2 align="center" ><%=descripcion%></H2>
</HEAD>
<script language="javascript">
function Regresar() {
FORM.submit 
return true
}
</script>
<BODY>
<div align="center">
<FORM NAME="FORM" ACTION="<%=destino%>"method="post">

<%
if ingresado = 5 then
%>
<table border="0" cellpadding="2" cellspacing="2">
<tr>
  <th nowrap>Pregunta: 
  </th>  
  <td nowrap><%=request.form("pregunta")%>
  </td>  
</tr>

<%
cont = Request.Form("contrespuesta")
for x = 1 to cont-1
%>
<tr>
<th nowrap>Respuesta: 
  </th>  
  <td nowrap><%=request.form("respuesta"&x)%>
  </td>  
</tr>  
<%
next
%>

</table>
<%
end if
%>
<input name="Aceptar" type="submit" value="Aceptar" onclick="return Regresar()">
</FORM>
</div>

</BODY>
</HTML>