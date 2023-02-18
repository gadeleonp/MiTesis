<%@ Language=VBScript %>
<%
nombre = Request.Form("nombre")
direccion = Request.Form("direccion")
telefono = Request.Form("telefono")
sitio = Request.Form("sitio")
email = Request.Form("email")
pais = Request.Form("pais")

dim db
dim resultado
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
sql = " insert into universidad (nombre,direccion,telefono,sitio,email,pais) values ('"
sql = sql + nombre + "','" + direccion + "','" + telefono + "','" + sitio + "','" + email +"',"+ pais + ")"
resultado = db.SQLOtras ("DBBAECYS",  sql)
Response.Redirect ("confirmacion.asp?ingresado=1")
%>
