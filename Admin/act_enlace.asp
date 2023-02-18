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
id = request.form("id")
 set db = server.createobject("SqlComandos.Comandos")
 sql = "update enlace set url = '"+url+"',texto = '"+texto+"',descripcion = '"+descripcion+"' where id = " & id
 

 resultado = db.sqlotras("DBBAECYS",sql)

response.redirect("actualizar_enlace.asp") 
%>
