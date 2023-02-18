<%@LANGUAGE="VBSCRIPT" %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
dim db
dim rs
dim sql 
dim id

id = request.form("id")

set db  = server.CreateObject("sqlcomandos.comandos")

sql = "select id,texto,url,descripcion from enlace where id = " & id

set rs = db.sqlselect("DBBAECYS",sql)

if not rs.eof then
  url = rs("url")
  texto = rs("texto")
  descripcion = rs("descripcion")
end if

set rs = nothing
set db = nothing

%>

<html>
<head>
<title>Actualizar enlace de interes.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<script language="javascript">
function version() {
if (ie) {
alert ("hola")
alert(navigator.appVersion.indexOf("4."))
}
else
alert("otro navegador")
////navigator.appVersion.indexOf("4.")==0
}
function Cancelar() 
{
document.FORM.action = "actualizar_enlace.asp"
document.FORM.submit()
}

function validacion_longitud(entered, min, max, alertbox){
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

function validacion() {
  	
  	
  	if(validacion_longitud(document.FORM.url,12,250,'El URL no puede ser menor de 12 caracteres ni mayor de 250')==false)
				return false
	if(validacion_longitud(document.FORM.texto,10,100,'El Texto no puede ser menor de 10 caracteres ni mayor de 100')==false)
				return false
	if(validacion_longitud(document.FORM.descripcion,10,400,'El URL no puede ser menor de 10 caracteres ni mayor de 400')==false)
				return false
					
  
  if (validacion_caracteres(document.FORM.url) == false)
  return false
  else
  if (validacion_caracteres(document.FORM.texto) == false)
  return false
  else
  if (validacion_caracteres(document.FORM.descripcion) == false)
  return false
  else
  return true
}

function validacion_caracteres(obj) 
{
  if (obj.value.indexOf("'")>=0) 
  {
  alert("Ingreso un caracter ( ' ) (no valido) en la casilla.");
  obj.select()
  obj.focus()
  return false
  }
  else
  if (obj.value.indexOf('"')>=0) 
  {
  alert("Ingreso un caracter ( \" ) (no valido) en la casilla.");
  obj.select()
  obj.focus()
  return false
  }
  else
  return true;
}
function Aceptar() {
	if (validacion() == true) 
	{
		document.FORM.action = "act_enlace.asp"
		document.FORM.submit()
	}
}
</script>
</head>
<body >
<div align="center">
<h2>Actualizar enlace de Interes</h2>
<form name="FORM" method="post" action="act_enlace.asp">
<table border="0" cellpadding="2" cellspacing="2" >
<tr>
	<th nowrap class="celdaBgBlueLeft"><div align="left">URL</div></th>
	<td class="celdaBgBlue05"><div align="left">
	  <input name="url" type="text" value="<%=url%>" size="100">
	  </div></td>
</tr>
<tr>
	<th nowrap class="celdaBgBlueLeft">&nbsp;
	  <div align="left">Texto a Mostrar&nbsp;</div></th>
	<td class="celdaBgBlue05"><div align="left">
	  <input name="texto" type="text" value="<%= texto %>" size="100" >
	  </div></td>
</tr>
<tr>
	<th nowrap class="celdaBgBlueLeft">&nbsp;
	  <div align="left">Breve descrpición del enlace&nbsp;</div></th>
	<td class="celdaBgBlue05"><div align="left">
	  
	  <textarea name="descripcion" cols="100" rows="4" ><%= descripcion %></textarea>
	  </div></td>
</tr>
</table><input name="id" type="hidden" value="<%= id %>">
</form>
<br>

<table border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td><input name="Aceptar" type="button" class="celdaBgBlue05" value="Aceptar" onclick="javascript:Aceptar()"></td>
    <td>&nbsp;</td>
    <td><input name="Cancelar" type="button" class="celdaBgBlue05" value="Cancelar" onclick="javascript:Cancelar()"></td>
  </tr>
</table>

</div>
</body>
</html>
