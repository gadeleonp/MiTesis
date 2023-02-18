<%@ Language=VBScript %>
<%
dim db
dim categoria
dim padre 
dim resultado

categoria = ""
padre = ""
categoria = Request.Form("categoria")
padre = Request.Form ("padre")

dim rs
set db = server.CreateObject("SQLComandos.Comandos")
sql = "select descripcion from categoria where descripcion = '"+categoria+"'"
set rs = db.SQLSelect("DBBAECYS",sql)
if rs.EOF <> true then
%>
<html>
<head>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</head>
<body>
<div align="center">
<h2>
Error: La categoria ingresada ya existe.
</h2>
<a href="ingreso_categoria.asp"> Regresar </a>
</div>
</body>
</html>
<%
set rs = nothing
set db = nothing

else
If InStr(1,padre,"0",1) = 0 Then
sql = " insert into categoria (descripcion,padre) values ('"+categoria+"',"+padre+")"
else
sql = " insert into categoria (descripcion,padre) values ('"+categoria+"',null)"
end if

resultado = db.SQLOtras("DBBAECYS",sql)
set rs = nothing
set db = nothing
Response.redirect("Confirmacion.asp?ingresado=3")
end if
%>