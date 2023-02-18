<%@ Language=VBScript %>
<HTML>
<HEAD>
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

<table aling="center" width="0%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td>Pregunta: </td><td><input name="pregunta" type="text"><td>
  </tr>
  <tr>
    <td>Numero de Respuestas</td>
	<td>
		<select name="respuestas" size="1">
		<option value="1" selected>Verdadero/Falso</option>
        <option value="2" >Dos</option>
        <option value="3">Tres</option>
      </select>
	<td>
  </tr>
  <tr>
    <td>Publicar</td><td>
	  <select name="publicar" size="1">
        <option value="1" selected>Si</option>
        <option value="2">No</option>
      </select></td>	
  </tr>
</table>
<br>
    <input name="Aceptar" type="submit" value="Aceptar">
    &nbsp;&nbsp; 
    <input name="Cancelar" type="button" value="Cancelar" onclick="javascript:Regresar()">
  </p>
</FORM>
</div>
</BODY>
</HTML>
