<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<head>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</head>
<BODY>
<DIV ALIGN="CENTER">

<TABLE BORDER="0" cellpadding="2" cellspacing="2">
  <TR>
    <TD class="celdaBgBlue">OPCIONES</TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./ingresar_universidad.asp">INGRESO DATOS UNIVERSIDAD</A></TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./eliminar_universidad.asp">ELIMINAR DATOS UNIVERSIDAD</A></TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./listado_universidad.asp">ACTUALIZAR DATOS UNIVERSIDAD</A></TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./mantusuarios.asp">REGRESAR AL MENU ANTERIOR</A></TD>
  </TR>
</TABLE>
</DIV>
</BODY>
</HTML>