<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script languege="javascript">
  function Reporte() {
  encuesta.submit()
  }
</script>

</HEAD>
<BODY onload="javascript:Reporte()">
<form name="encuesta" method="post" action="confirmacion.asp?Ingresado=5">
<input name="Ingresado" value="5" type="hidden">
<%
dim db
dim resultado
dim rs
set db = server.CreateObject("SQLComandos.Comandos")
dim fecha
fecha = now

sql = "insert into pregunta (descripcion, numero_res,publicar,fecha) values  ('" & session("pregunta") & "'," & session("respuestas") & "," & session("publicar") &",'"& now & "')"
resultado = db.SQLOtras ("DBBAECYS",sql)
 
sql = "select id from pregunta where descripcion like '%" &session("pregunta") & "%' and fecha = '"& fecha &"'"

set rs = db.SQLSelect ("DBBAECYS",sql)
if not rs.eof then
  pregunta = rs("id")
end if
set rs = nothing

tipo_respuesta = Request.QueryString("tipo_respuesta")

if tipo_respuesta = 1 then
 sql = "update pregunta set numero_res = 2 where id = " & pregunta
 resultado = db.SQLOtras("DBBAECYS",sql)
 sql = "insert into opciones_respuesta (descripcion,total,pregunta) values ('VERDADERO',0," & pregunta &")"
 resultado = db.SQLOtras("DBBAECYS",sql)
 sql = "insert into opciones_respuesta (descripcion,total,pregunta) values ('FALSO',0," & pregunta &")"
 resultado = db.SQLOtras("DBBAECYS",sql)
	%>
	<input type="hidden" name="respuesta1" value="<%="VERDADERO"%>">
	<input type="hidden" name="respuesta2" value="<%="FALSO"%>">	
	<%
else
 for cont = 1 to session("respuestas") 
 		texto_respuesta = "respuesta"&cont
 		 
		Respuesta = Request.form(texto_respuesta) 
		sql = "insert into opciones_respuesta (descripcion,total,pregunta) values ('"& respuesta & "',0," & pregunta &")"
		resultado = db.SQLOtras("DBBAECYS",sql)
		%>
		<input type="hidden" name="respuesta<%=cont%>" value="<%=respuesta%>">
		<%
  next
end if
 %>
 
<%
set db = nothing
%>
<input type="hidden" name="contrespuesta" value="<%if (tipo_respuesta = 1) then Response.write "3" else Response.write cont%>">
<input type="hidden" name="pregunta" value="<%=session("pregunta")%>">
</BODY>
</HTML>
