<%@LANGUAGE="VBSCRIPT" %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Configuracion Correos</title>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<table width="200" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <td class="celdaBgBlue">OPCIONES</td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="./ConfCorreoNoti.asp">Correo Noticias</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="./ConfCorreoForo.asp">Correo Foros</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="./ConfCorreoEmp.asp">Correo Empresas</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="./opciones.asp">Regresar</a></td>
  </tr>
</table>

</body>
</html>
