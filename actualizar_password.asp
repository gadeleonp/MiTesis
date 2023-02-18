<%@LANGUAGE="VBSCRIPT"%>
<%
if isempty(session("id")) then
    Response.Redirect("login.asp")
end if
%>
<html>
<head>
<title>Actualizar Contraseña</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="styles/interior.css" rel="stylesheet" type="text/css">
<script language="javascript">
function Regresar() {
  document.location = "opciones_datos.asp"
}

function validacion_password(entered, min, max, alertbox){
	with (entered){
		if ((value.length<min) || (value.length>max)){
			if (alertbox!="") {
				alert(alertbox);
			} 
			select();
			focus();
			return false;
		}
		else {
			return true;
		}
	}
}

 function verificar(formaingreso) 
  {
  if (formaingreso.action == "actualiza_password.asp") 
  {
 
			if(validacion_password(formaingreso.password,7,15,'El password debe tener entre 7 y 15 caracteres')==false){
				return false;
			}

			if(validacion_password(formaingreso.passwordnueva,7,15,'El password debe tener entre 7 y 15 caracteres')==false){
				return false;
			}

			if(validacion_password(formaingreso.confirmapassword,7,15,'La confirmación del password debe tener entre 7 y 15 caracteres')==false){
				return false;
			}
		
			if (formaingreso.passwordnueva.value != formaingreso.confirmapassword.value) 
			{
				alert("Las contraseñas son diferentes.\nPor favor vuelva a ingresarlas.")
				formaingreso.passwordnueva.focus()
				formaingreso.passwordnueva.select()
				//formaingreso.confirmapasswor.select()
				return false;
			}
	}
	  return true;
	}
</script>
</head>
<body onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}">
<div align="center">
<h2>Cambiar Contraseña</h2>
<form onsubmit="return verificar(this)" name="Form" action="actualiza_password.asp" method="post">
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/spacer.gif" width="10" height="10"></td>
        </tr>
        <tr>
          <td><table width="132" height="110" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td height="24"><table width="200" height="24" border="0" align="center" cellpadding="0" cellspacing="0" background="images/bar02.gif">
                    <tr>
                      <td width="9">&nbsp;</td>
                      <td width="776"><div align="left"><strong><font size="2">Log&iacute;n</font></strong></div></td>
                      <td width="10" align="right">&nbsp;</td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="92" align="center"><form name="FORM1" method="post" action="./verifica_usuario.asp">
                    <table width="95%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="left"><strong><font color="#000099" size="2">Contrase&ntilde;a  Anterior:</font></strong></td>
                      </tr>
                      <tr>
                        <td align="center"><input name="password" type="password" id="password" size="15"></td>
                      </tr>
                      <tr>
                        <td align="left"><font color="#000099" size="2"><strong>Nueva contrase&ntilde;a:</strong></font></td>
                      </tr>
                      <tr>
                        <td align="center"><input name="passwordnueva"type="password"  size="15"></td>
                      </tr>
					                        <tr>
                        <td align="left"><font color="#000099" size="2"><strong>Reingrese la nueva contrase&ntilde;a:</strong></font></td>
                      </tr>
                      <tr>
                        <td align="center"><input name="confirmapassword" type="password" size="15"></td>
                      </tr>
                    </table>
                    <br>
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td>
<input name="Aceptar" type="submit" class="celdaBgBlue05" value="Aceptar">
</td>
<td width="17"></td>
<td><input name="Cancelar" type="button" class="celdaBgBlue05" value="Cancelar" onclick="javascript:Regresar()"></td>
</tr>

</table>
</form>
</div>
</body>
</html>
