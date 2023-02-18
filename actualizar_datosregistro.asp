<%@ Language=VBScript %>

<%
  if session("id") = 0 then
	response.redirect("Login.asp")
   end if
%>
<%
dim db
Dim rs
dim resultado
set db = server.CreateObject ("SQLComandos.Comandos")

Dim nombre,apellido,login,password,direccion,email,registro,telefono

universidad = ""

id = Request.Form("id")
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
pais = Request.Form("nacionalidad")
documento = Request.Form("documento")
tipo_id = Request.Form("tipo_id")
sexo = request.form("sexo")

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
'sql = "update usuario set nombre = '" & nombre &"', apellido = '"& apellido & "', login ='" & login & "', password = '" & password & "', direccion = '" &direccion & "', telefono = '" & telefono & "', mail = '" & email & "', universidad = " & universidad &", pais = "& pais & ", carnet = '" & carnet & "', colegiado = '" & colegiado & "', tipo = " & tipo_id & ", sexo = "& sexo &", recibir_mail = '" & recibir_correo & "',tipo_mail ='"&conf_correos & "' where id = " & id
sql = "update usuario set nombre = '" & nombre &"', apellido = '"& apellido & "', login ='" & login & "',  direccion = '" &direccion & "', telefono = '" & telefono & "', mail = '" & email & "', universidad = " & universidad &", pais = "& pais & ", carnet = '" & carnet & "', colegiado = '" & colegiado & "', tipo = " & tipo_id & ", sexo = "& sexo &", recibir_mail = '" & recibir_correo & "',tipo_mail ='"&conf_correos & "' where id = " & id
resultado = db.SQLOtras("DBBAECYS",sql)


%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language="javascript">
 function Continuar(Forma) {
  Forma.submit()
 return true
 }
</script>
<BODY onload="return Continuar(forma2)">
<form name="forma2" method = "post" action="actualiza_preferencias.asp">
<input name="login" type="hidden" value ="<%=login%>">
</form>
</BODY>
</HTML>

<%
set db = nothing
%>
