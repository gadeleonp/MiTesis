<%@ Language=VBScript %>
<%
dim db
dim pais
pais = Request.Form("pais")
dim rs
dim resultado
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select nombre from pais where nombre = '"+pais+ "'"
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
%>
<h2>El Pais ya fue ingresado previamente.</h2>
<a href="./mant_paises.asp">Regresar...</a>
<%
else
sql = " insert into pais (nombre) values ('" + pais + "')"
resultado = db.SQLOtras("DBBAECYS",sql)
set rs = nothing
Response.Redirect ("Confirmacion.asp?Ingresado=2")
end if
set rs = nothing
%>
