<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim sql
dim rs

url = request.form("url")
texto = request.form("texto")
descripcion = request.form("descripcion")

 set db = server.createobject("SqlComandos.Comandos")
 sql = "insert into enlace (url,texto,descripcion) values ('"+url+"','"+texto+"','"+descripcion+"')"
 
 resultado = db.sqlotras("DBBAECYS",sql)

response.redirect("listado_url.asp") 
%>
