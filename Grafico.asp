<%@ Language=VBScript %>
<%
dim db
dim cod_pregunta,pregunta
cod_pregunta = Request.QueryString("pregunta")

dim rs
set db = server.CreateObject ("SQLComandos.Comandos")

dim descripcion(10) 
dim total(10) 
   
sql ="select descripcion,numero_Res from pregunta where id = "& cod_pregunta
set rs = db.SQLSelect ("DBBAECYS",sql)

if  not rs.EOF  then
	pregunta = Rs("descripcion")
	numero_Res = rs("numero_Res")
end if
set rs = nothing



%>
<html>
<head>
<!--<link href="./theme/Master.css" rel="stylesheet" type="text/css">-->
<link href="./styles/interior.css" rel="stylesheet" type="text/css">
<title>Encuesta</title>
<style type="text/css">
<!--
td {
	letter-spacing: normal;
	text-align: center;
	vertical-align: middle;
	word-spacing: normal;
}
-->
</style>
<style type="text/css">
<!--
.tdimage {
	background-attachment: fixed;
	background-image: url(grafics.asp);
	background-repeat: no-repeat;
	background-position: center center;
}
-->
</style>
</head>
<script language="javascript">
function cerrarventana() {
	   window.close()
 	}
</script>

<body>
<%
sql = "select count(id) num_resp from opciones_respuesta where pregunta = " & Cod_pregunta & " order by total desc "



sql = "select id,descripcion,total from opciones_respuesta where pregunta = "&cod_pregunta & " order by total desc "

set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof  then
cont = 0

 while rs.EOF <> true
   descripcion_respuesta = rs("descripcion")
   
    descripcion(cont) = rs("descripcion")
    total(cont) = rs("total")   
cont = cont + 1

session("cont") = cont

rs.movenext
wend    
end if
set rs = nothing

session("total") = total
session("descripcion") = descripcion

%>
<table align="center" border="0" cellpadding="2" cellspacing="2">
<tbody>
<tr>
<th><h4><%=pregunta%></h4></th>
</tr>
<tr>
<td height="3"><img src="grafics.asp" ></td>
</tr>
<td>
<input name="Cerrar" type="button" class="celdaBgBlue05" onclick="javascript:cerrarventana()" value="Cerrar">
</td>
<tr>

</tr>
</tbody>
</table>
</body>
</html>
