<%@LANGUAGE="VBSCRIPT" %>
<%

 dim usuario
usuario = ""
usuario = session("id")
if len(usuario) = 0 then
  response.redirect("login.asp")
end if

dim password
dim password2
dim confirma

password = ""
password2= ""
confirma = ""

password = request.form("password")
password2 = request.form("passwordnueva")
confirma = request.form("confirmapassword")
dim sql
dim db
dim rs

set db =  server.createobject("SQLComandos.Comandos")
 sql = "select password from usuario where id = " & usuario & " and password = '" & password & "'" 


set rs = db.SQLSelect("DBBAECYS",sql)
noexiste = 0
if not rs.eof then
  password_dbb = rs("password")
else
 noexiste = 1
end if


if password = password_dbb then
else
Response.Redirect("actualizar_password.asp?Error=1")
end if


set rs = nothing
 if noexiste = 1 then
 else
 sql = "update usuario set password = '" & password2 & "' where id = " & usuario & " and password = '" &password & "'"
 end if
 

 db.SQLOtras "DBBAECYS",sql
Response.redirect("opciones_datos.asp")
%>