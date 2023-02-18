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
<script language="JavaScript">
function validar(form) {
  if (form.pregunta.value.length== 0) {
	alert ("Error: No ha ingresado la respuesta.");
	forma.pregunta.focus()
	return false
  }
}
function Regresar() {
  window.location="mantencuesta.asp"
}
</script>

<BODY>
<div align="center">
<FORM action="ingresar_pregunta.asp" method=POST id=form1 name="forma" onsubmit= "return validar(this)">
<h1>Creación de Preguntas</h1>
<table aling="center" width="0%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td class="celdaBgBlueLeft">Pregunta: </td>
    <td class="celdaBgBlue05"><input name="pregunta" type="text"><td class="celdaBgBlue05">
  </tr>
  <tr>
    <td class="celdaBgBlueLeft">Numero de Respuestas</td>
	<td class="celdaBgBlue05">
		<select name="respuestas" size="1">
		<option value="1" selected>Verdadero/Falso</option>
        <option value="2" >Dos</option>
        <option value="3">Tres</option>
        <option value="4">Cuatro</option>
        <option value="5">Cinco</option>
        <option value="6">Seis</option>
        <option value="7">Siete</option>
      </select>
	<td class="celdaBgBlue05">
  </tr>
  <tr>
    <td class="celdaBgBlueLeft">Publicar</td>
    <td class="celdaBgBlue05">
	  <select name="publicar" size="1">
        <option value="1" selected>Si</option>
        <option value="2">No</option>
      </select></td>	
  </tr>
</table>
<br>
    <input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar">
    &nbsp;&nbsp; 
    <input name="Cancelar" type="button" class="celdaBgBlue05" onclick="javascript:Regresar()" value="Cancelar">
  </p>
</FORM>
</div>
</BODY>
</HTML>
