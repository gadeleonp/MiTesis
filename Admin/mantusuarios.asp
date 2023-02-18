<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<title>Mantenimiento de Usuarios</title>
</HEAD>

<BODY>

<TABLE WIDTH=31% BORDER=0 CELLSPACING=2 CELLPADDING=2 align="center">
  <TR> 
    <th nowrap class="celdaBgBlue">
     Mantenimiento de Opciones para 
        usuarios
    </th>
  </TR>
  <TR> 
    <TD class="celdaBgBlue05"> 
      <div align="center"><a href="mant_universidades.asp">Mantenimiento Universidades</a></div>    </TD>
  </TR>
  <TR> 
    <TD class="celdaBgBlue05"> 
      <div align="center"><a href="mant_paises.asp">Mantenimiento Paises</a></div>    </TD>
  </TR>
  <TR> 
    <TD class="celdaBgBlue05"> 
      <div align="center"><a href="mant_areas.asp">Mantenimiento Areas-Subareas 
    de conocimiento</a></div>    </TD>
  </TR>
  <TR> 
    <TD class="celdaBgBlue05"> 
      <div align="center"><a href="opciones.asp">Regrear al menu anterior</a></div>    </TD>
  </TR>
</TABLE>

</BODY>
</HTML>
