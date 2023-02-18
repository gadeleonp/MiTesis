<%@ Language=VBScript %>
<%
dim db
dim area
dim categoria 
dim explicacion

area = ""
padre = ""
explicacion = ""
area = Request.Form("area")
categoria = Request.Form ("categoria")
explicacion = Request.Form ("explicacion")

dim rs
set db = server.CreateObject ("SQLComandos.Comandos")
sql = "select descripcion from area where descripcion = '"+area+"'" 
'sql = "select descripcion from area where descripcion = '" + area + "' and categoria in (select id from categoria where padre is  "
set rs = db.SQLSelect("DBBAECYS",sql)
if rs.EOF <> true then
%>
<html>
<head >
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</head>
<body>
<div align="center">
<h2>
Error: El area no pudo ser ingresada.
</h2>
<a href="ingreso_area.asp">Regresar</a>
</div>
</body>
</html>
<%
set rs = nothing
set db = nothing
else
sql = " insert into area (descripcion,categoria,explicacion) values ('"+area+"',"+categoria+",'"+explicacion+"')"
resultado = db.SQLOtras("DBBAECYS",sql)
set rs = nothing
set db = nothing
Response.redirect("Confirmacion.asp?ingresado=4")
end if
%>