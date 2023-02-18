<%@ Language=VBScript %>
<html>
<head>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
dim db
dim universidades
dim rs
dim resultado
set db = server.CreateObject ("SQLComandos.Comandos")
universidades = Request.Form("seleccion")
sql = "delete from universidad where id in (" & universidades & ") and id not in ( select distinct(universidad) from usuario )"
resultado = db.SQLOtras ("DBBAECYS", sql)

sql = "select distinct(u.universidad),p.nombre pais, uv.nombre from usuario u,universidad uv, pais p where u.universidad in (" & universidades & ")  and uv.id = u.universidad and uv.pais = p.id"
 
set rs = db.sqlselect("DBBAECYS", sql)

if not rs.EOF then
while rs.EOF <> true 
%>
 <h2>La universidad "<%= rs("nombre") %>" de <%=rs("pais")%> no fue eliminada por que esta asignada a un usuario </h2><br>
<%
rs.MoveNext 
wend
%>
<a href="eliminar_universidad.asp">Regresar</a>
<%
else
Response.Redirect("eliminar_universidad.asp?error=1")
end if
set rs = nothing
%>
</body>
</html>




