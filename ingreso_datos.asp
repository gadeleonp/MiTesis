<%@ Language=VBScript %>
<!--#INCLUDE FILE="genera_clave.asp"-->
<%

session.Abandon()
dim db

Dim rs
dim resultado
set db = server.CreateObject ("SQLComandos.Comandos")

Dim nombre,apellido,login,password,direccion,email,registro,telefono,sexo,nacionalidad,pais

universidad = ""


nombre = Request.Form("nombre")
apellido = Request.Form("apellido")
login = Request.Form("login")
'password = Request.Form("password")
direccion = Request.Form("direccion")
telefono = Request.Form("telefono")
email = Request.Form("email")
tipo = Request.Form("tipo_id")
documento = Request.Form("documento")
universidad = Request.Form("universidad")
pais = Request.Form("pais")
documento = Request.Form("documento")
tipo_id = Request.Form("tipo_id")
nacionalidad = request.form("nacionalidad")
sexo = request.form("sexo")
carrera = request.form("carrera")
carrera_esp = request.form("carrera_esp")

recibir_correo = request.form("correo") 'Si el usuario desea recibir un mail este tendra valor 1
if len(recibir_correo)  < 2 then
recibir_correo = "SI"
end if

correo_foros = request.form("correo_foros") 	 		'Recibir mail de foros
correo_noticias = request.form("correo_noticias")		'  "     "     "  noticias
correo_contactos = request.form("correo_contactos")		'  "     "     "  contactos

conf_correos = trim(correo_foros) & trim(correo_noticias) & trim(correo_contactos) '  Configuracion de correos

if cint(tipo_id) = 1 then
	carnet = documento
else
	if cint(tipo_id) = 2 then
	colegiado = documento
	end if
end if

uni = cint(universidad)

if uni = -1 then
universidad = "0"
end if

if len(universidad) > 0 or universidad <> ""  then
else
universidad = "0"
end if
sql = "select login from usuario where login = '" & login &"'"
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
<HTML>
<HEAD>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language = "javascript">
 function regresar(Forma) {
  Forma.action = "ingdatospersonales.asp"
  Forma.submit()
  return true 
 }
</script>
<BODY onload="return regresar(forma)">
<h2>Error el codigo de usuario o login ya existe favor volver a ingresarlo.</h2>
<form name="forma" action ="ingdatospersonales.asp" method = "post">
<input name="nombre" type ="hidden" value = "<%=nombre%>">
<input name="login" type ="hidden" value = "<%=login%>">
<input name="apellido" type ="hidden" value = "<%=apellido%>">
<input name="direccion" type ="hidden" value = "<%=direccion%>">
<input name="telefono" type ="hidden" value = "<%=telefono%>">
<input name="email" type ="hidden" value = "<%=email%>">
<input name="universidad" type ="hidden" value = "<%=universidad%>">
<input name="documento" type ="hidden" value = "<%=documento%>">
<input name="tipo_id" type ="hidden" value = "<%=tipo_id%>">
<input name="pais" type ="hidden" value = "<%=pais%>">
<input name="sexo" type ="hidden" value = "<%=nacionalidad%>">
<input name="error" type ="hidden" value = "1">
</form>
</BODY>
</HTML>

<%
else

'sql = "insert into usuario (nombre,apellido,login,password,direccion,telefono,mail,universidad,pais,carnet,colegiado,tipo,sexo,ultima_visita,recibir_mail,tipo_mail) values ('" & nombre &"','"& apellido & "','" & login & "','" & password & "','" &direccion & "','" & telefono & "','" & email & "'," & universidad &","& nacionalidad & ",'" & carnet & "','" & colegiado & "'," & tipo_id & "," & sexo & ",'" & now & "','" & recibir_correo & "','" & conf_correos &"'"  & ")"
password = clave()
sql = "insert into usuario (nombre,apellido,login,password,direccion,telefono,mail,universidad,pais,carnet,colegiado,tipo,sexo,ultima_visita,recibir_mail,tipo_mail,habilitada,carrera,carrera_esp) values ('" & nombre &"','"& apellido & "','" & login & "','" & password & "','" &direccion & "','" & telefono & "','" & email & "'," & universidad &","& nacionalidad & ",'" & carnet & "','" & colegiado & "'," & tipo_id & "," & sexo & ",'" & now & "','" & recibir_correo & "','" & conf_correos &"'"  & ",1,"&carrera&",'"&carrera_esp&"')"

resultado = db.SQLOtras("DBBAECYS",sql)

end if
set rs = nothing
%>
<HTML>
<HEAD>
<META NAME="GENERATOR">
</HEAD>
<script language="javascript">
 function Continuar(Forma) {
  Forma.submit()
 return true
 }
</script>
<BODY onload="return Continuar(forma2)">
<form name="forma2" method = "post" action="ingreso_preferencias.asp">
<input name="login" type="hidden" value ="<%=login%>">
</form>
</BODY>
</HTML>

<%
set db = nothing
%>
