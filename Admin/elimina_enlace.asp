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

id = request.form("id")


 set db = server.createobject("SqlComandos.Comandos")
 sql = "delete from enlace where id = "& id
 
 resultado = db.sqlotras("DBBAECYS",sql)

response.redirect("eliminar_enlace.asp") 
%>
