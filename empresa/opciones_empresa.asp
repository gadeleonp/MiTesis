<%@ Language=VBScript %>
<%
dim user
user = session("UsEmpresa")
if isempty(user) then
  Response.Redirect("../Login.asp")
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<title>Opciones Empresariales</title>
<script language="javascript">
 function Enlace(Opcion) 
 {
	 if (Opcion=="1") 
	 {
		destino = "buscar_usuario.asp"
	 }
	 if (Opcion=="2") {
 		destino = "../salir.asp"
	 }
	 document.window.location = destino
}
 </script>
</HEAD>
<BODY>
<div align="center" >
<h2>Opciones</h2>
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td class="celdaBgBlue05">
<a href="./buscar_usuario.asp"  onClick="javascript:Enlace('1')">Busquedas de usuarios</a>
</td>
</tr>
<tr>
<td class="celdaBgBlue05">
<a href="../salir.asp" onClick="javascript:Enlace('1')">Terminar Sesion</a>
</td>
</tr>
</table>
</div>
</BODY>
</HTML>
