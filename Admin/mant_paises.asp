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
<BODY>
<DIV ALIGN="CENTER">

<TABLE BORDER="0" cellpadding="2" cellspacing="2">
  <TR>
    <TD class="celdaBgBlue">OPCIONES</TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./ingresar_pais.asp">INGRESO PAIS</A></TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./eliminar_pais.asp">ELIMINAR PAIS</A></TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./listado_pais.asp">ACTUALIZAR DATOS PAIS</A></TD>
  </TR>
  <TR>
  <TD class="celdaBgBlue05"><A href="./mantusuarios.asp">REGRESAR AL MENU ANTERIOR</A></TD>
  </TR>
</TABLE>
</DIV>
</BODY>
</HTML>
