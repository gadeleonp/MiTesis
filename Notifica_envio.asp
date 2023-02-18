<%@ Language=VBScript %>
<%
dim id
id = request.form("id")
destino = request.form("mail")
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
  function envio() {
      document.Form.submit()
  	 return true;
  }
</script>
</HEAD>
<BODY>
<div align="center">
<h1>Notificación</h1>
</div>
<h2>
Su contraseña ha sido enviada a la cuenta de correo que ingreso previamente.
Revise su buzón de correo y copia la contraseña para poder ingresar al sistema. Le recomendamos que la modifique
despues de ingresar.
<h2>

<br>
<h2>Si no recibio el correo electrónico por favor presione <a href="javascript:envio()" ><font size="4">aqui</font></a></h2>
<h2>
Para continuar con su configuración ingrese con numero de cuenta y contraseña nuevamente.
 <a href="./Login.asp" ><font size="4">aqui</font></a>
<h2>
<form name="Form" action="enviar_clave.asp" method="post">
<input name="mail" value="<%=destino%>" type="hidden">
<input name="id" value="<%=id%>" type="hidden">
</form>
</BODY>
</HTML>
