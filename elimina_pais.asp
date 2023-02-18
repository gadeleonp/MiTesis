<%@ Language=VBScript %>
<html>
<body>
<%
dim db
dim resultado
dim paises
dim rs
set db = server.CreateObject ("SQLComandos.Comandos")

paises = Request.Form("seleccion")
sql = "delete from pais where id in (" & paises & ") and id not in ( select distinct(pais) from usuario )"

resultado = db.SQLOtras("DBBAECYS",sql)

sql = "select distinct(u.pais), p.nombre pais from usuario u, pais p where u.pais in (" & paises & ")  and u.pais = p.id"

set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.eof then
while rs.eof <> true 
%>
 <h2>El pais "<%= rs("pais") %>"  no fue eliminado por que esta asignada a un usuario </h2><br>
<%
rs.MoveNext 
wend
%>
<a href="eliminar_pais.asp">Regresar</a>
<%
set rs = nothing
set db = nothing
else
set rs = nothing
set db = nothing
Response.Redirect("eliminar_pais.asp?error=1")
end if

%>
</body>
</html>




