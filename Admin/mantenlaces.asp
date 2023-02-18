<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<html>
<head>
<title>Mantenimiento de Enlaces</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function Regresar() {
	  document.location = "opciones.asp"
  }
  function Crear() {
    document.location = "crear_enlace.asp"
  }
  function Eliminar() {
    document.location = "eliminar_enlace.asp"
  }
  function Actualizar() {
      document.location = "actualizar_enlace.asp"

  }
  function Reporte() {
    document.location = "listado_url.asp"
  }
 </script>
</head>
<body>
<table border="0" align="center" cellpadding="2" cellspacing="2">
  <tr>
    <th class="celdaBgBlue" scope="col">Opciones</th>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="javascript:Crear()">Crear Enlace</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="javascript:Eliminar()">Eliminar Enlace</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="javascript:Actualizar()">Actualizar Enlace</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="javascript:Reporte()">Reporte de Enlaces</a></td>
  </tr>
  <tr>
    <td class="celdaBgBlue05"><a href="javascript:Regresar()">Regresar</a></td>
  </tr>
</table>

</body>
</html>
