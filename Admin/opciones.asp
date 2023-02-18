<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<title>Opciones Administrador</title>
<BODY>
<table width="17%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr height="10" > 
    <th height="10" nowrap class="celdaBgBlue"> 
        Opciones
    </th>
  </tr>
  <tr> 
    <td height="23" nowrap class="celdaBgBlue05"> 
	  <div align="left"><a href="mantnoticias.asp" target= "_top" class="celdaBgBlue05">Mantenimiento Noticias</a>
	    </div></td>
  </tr>
  <tr> 
    <td nowrap class="celdaBgBlue05"> 
		<div align="left"><a href="mantusuarios.asp" target= "_top">Mantenimiento Opciones para Usuarios</a> </div></td>
  </tr>
  <tr> 
    <td nowrap class="celdaBgBlue05"> 
		<div align="left"><a href="./mantforos.asp">Mantenimiento Foros</a>
	      </div></td>
  </tr>
    <tr> 
    <td nowrap class="celdaBgBlue05"> 
		<div align="left"><a href="./configuracion_correos.asp" target= "_top">Mantenimiento Correos</a> </div></td>
    </tr>
      <tr>
      <td nowrap class="celdaBgBlue05">
		<div align="left"><a href="mantencuesta.asp" target= "_top">Mantenimiento Encuestas</a>
	      </div></td>
    </tr>
	  <tr> 
    <td nowrap class="celdaBgBlue05"> 
		<div align="left"><a href="mantenlaces.asp" target= "_top">Mantenimiento de Enlaces de Interes</a> </div></td>
  </tr>
  <tr> 
    <td nowrap class="celdaBgBlue05"> 
		<div align="center"><a href="login.asp" target= "_top">Salir</a> </div></td>
  </tr>
</table>
</BODY>
</HTML>
