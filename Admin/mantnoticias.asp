<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY>
<h1 align="center">Mantenimiento de Noticias</h1>

<table width="14%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr> 
    <td class="celdaBgBlue"> 
      <div align="center"><b>Opciones</b></div>    </td>
  </tr>
  <tr> 
    <td class="celdaBgBlue05"> 
      <div align="center"><a href="ingresonoticia.asp">Agregar Noticia</a></div>    </td>
  </tr>
  <tr> 
    <td class="celdaBgBlue05"> 
      <div align="center"><a href="actualizanoticia.asp">Modificiar Notica</a></div>    </td>
  </tr>
  <tr> 
    <td class="celdaBgBlue05"> 
      <div align="center"><a href="eliminarnoticia.asp">Eliminar Noticia</a></div>    </td>
  </tr>
  <tr> 
    <td class="celdaBgBlue05"> 
      <div align="center"><a href="listadonoticia.asp">Reporte de Noticias</a></div>    </td>
  </tr>
	<tr>
	 <td class="celdaBgBlue05"> 
      <div align="center"><a href="opciones.asp">Regresar al menu anterior</a></div>    </td>
	</tr>
</table>
</BODY>
</HTML>
