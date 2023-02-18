<html>
<head>
<style type="text/css">
<!--
.style1 {color: #0099FF}
-->
</style>
<link href="./styles/interior.css" rel="stylesheet" type="text/css">

</head>
<script language="javascript">

function Destino(Opcion) {
if (Opcion == "1") {
destino = "./principal.asp?ver=0&path=ingdatospersonales.asp"
window.parent.location = destino
}
if (Opcion == "2") {
destino = "./principal.asp?ver=0&path=Login.asp"
window.parent.location = destino
}
if (Opcion == "3") {
destino = "./principal.asp?ver=0&path=noticias.asp"
window.parent.location = destino
}

if (Opcion == "4") {
destino = "./admin/Login.asp"
window.parent.location  = destino
//parent.principal.location = destino
}
}
</script>
<body >
<table width="365" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="791" colspan="2"><img src="images/logo01.gif" width="323" height="47"></td>
  </tr>
  <tr>
    <td rowspan="2" valign="top"><img src="images/logo02.gif" width="44" height="29"></td>
    <td valign="top"><img src="images/logo03.gif" width="321" height="14"></td>
  </tr>
  <tr>
  <%
if not isempty(session("id")) then
%>
<td><div align="center"><strong><font color="#000099" size="1">Usuario: <%=session("nombre")%></font></strong></div></td>
<%	
end if
%>
  </tr>
</table>
</html>
