<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<div align="center">
<h2>Mantenimiento de Areas</h2>
  <table border="0" cellpadding="2" cellspacing="2" aling="center" width="21">
    <tr>
      <td nowrap class="celdaBgBlue">OPCIONES</td>
    </tr>
    <tr> 
      <td nowrap class="celdaBgBlue05"> 
      <div align="center"><a href="./ingreso_categoria.asp">Ingresar Categoria</a></div>      </td>
    </tr>
    <tr> 
      <td nowrap class="celdaBgBlue05"> 
      <div align="center"><a href="./ingreso_area.asp">Ingresar Area</a></div>      </td>
    </tr>
    <tr> 
      <td nowrap class="celdaBgBlue05"> 
      <div align="center"><a href="./mantusuarios.asp">Regresar al menu anterior</a>.</div>      </td>
    </tr>
  </table>
</div>
</BODY>
</HTML>
