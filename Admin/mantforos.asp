<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<META>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<table width="200" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <th nowrap><div align="center" class="celdaBgBlue">Opciones</div></th>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><div align="center"> <a href="./crearforo.asp" title="Crear Foro" target="_self">Crear Foro</a></div></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><div align="center">  <a href="./eliminarforo.asp" title="Eliminar Foro" target="_self">Eliminar Foro</a></div></div></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><div align="center">  <a href="./opciones.asp" title="Regresar" target="_self">Regresar</a></div></div></td>
  </tr>
</table>

</BODY>
</HTML>
