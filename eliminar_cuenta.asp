<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
if isempty(session("id")) then
    Response.Redirect("login.asp")
end if
dim error1
%>
<%
  error1 = request.QueryString("error")

 dim db
 dim sql
  set db = server.CreateObject("SqlComandos.Comandos")
  sql = "select nombre,apellido,mail, login from usuario where id = " & session("id")
  
  set rs = db.sqlselect("DBBAECYS",sql)
  
  if not rs.eof then
    nombre = rs("nombre") & " "& rs("apellido")
	mail = rs("mail")
	login =rs("login")
  end if
  
%>
<html>
<head>
<title>Eliminar cuenta</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript">
	function Desactivar() 
	{
	var accion;
	 accion = confirm("Esta opción desactivara su cuenta permanentemente. \n ¿Desea continuar?");   
	 if (accion == true)
		{
			if (verificar(formaingreso))
			formaingreso.submit()
		}
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
  if (formaingreso.action == "elimina_cuenta.asp") 
  {
			if(validacion_password(formaingreso.password1,7,15,'El password debe tener entre 7 y 15 caracteres')==false){
				return false;
			}
			if(validacion_password(formaingreso.password2,7,15,'La confirmación del password debe tener entre 7 y 15 caracteres')==false){
				return false;
			}
		
			if (formaingreso.password1.value != formaingreso.password2.value) 
			{
				alert("Las contraseñas son diferentes.\nPor favor vuelva a ingresarlas.")
				formaingreso.password2.value = ""
				formaingreso.password1.focus()
				formaingreso.password1.select()
				return false;
			}
	}
	  return true;
	}
  </script>
  <script language="javascript">
    function Cancelar() 
	  {
		  alert(window.location)
		  window.location = "opciones_datos.asp"
	  }
  </script>
<link href="styles/interior.css" rel="stylesheet" type="text/css">

</head>
<body>
<h2 align="center">Eliminación de Cuenta</h2>
<br>
<%if error1 = 1 then%><h4 align="center"><span style="font-size:16px;color:#FF0000; text-align:center">Contraseña Invalida</span></h4>
<%else%>
<br>
<%end if%>
<form name="formaingreso" method="post" action="elimina_cuenta.asp" >
<table border="0" align="center" cellpadding="2" cellspacing="2" width="100%">
  <tr class="reglon01">
    <th width="25%" class="celdaBgBlueLeft" scope="row">Nombre Completo:</th>
    <td width="50%"class="celdaBgBlue05"><%=Nombre%></td>
  </tr>
  <tr class="reglon02">
    <th class="celdaBgBlueLeft" scope="row">Cuenta: </th>
    <td class="celdaBgBlue05"><%=Login%></td>
  </tr>
  <tr class="reglon01">
    <th class="celdaBgBlueLeft" scope="row">Correo:</th>
    <td class="celdaBgBlue05"><%=mail%></td>
  </tr>
  <tr class="reglon02">
    <th class="celdaBgBlueLeft" scope="row">Ingreso Contraseña:</th>
    <td class="celdaBgBlue05"><input name="password1" type="password" maxlength="15" size="20">
		
	</td>
  </tr>
  <tr class="reglon01">
    <th class="celdaBgBlueLeft" scope="row">Reingreso Contraseña:</th>
    <td class="celdaBgBlue05"><input name="password2" type="password" maxlength="15" size="20" ></td>
  </tr>
</table>

  <br>
<table align="center" border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="Aceptar" type="button" class="celdaBgBlue05" onclick="javascript:Desactivar()" value="Aceptar"></td>
<td width="17"></td>
<td><input name="Cancelar" type="button" class="celdaBgBlue05" onclick="window.location='opciones_datos.asp'" value="Cancelar"></td>
</tr>
</table>
</form>
</body>
</html>
