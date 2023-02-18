<%@ Language=VBScript %>
<%
dim db
dim direccion
dim correlativo
dim documento
dim noticia
dim resultado

set db = server.CreateObject("SQLComandos.Comandos")

direccion = Request.QueryString ("direccion")
correlativo = Request.querystring ("correlativo")
documento = Request.QueryString ("documento")
noticia = Request.QueryString ("noticia")

sql = "update documento_noticia "

if direccion = "2" then 'abajo
sql = sql + " set correlativo = correlativo + 1 where noticia = "& noticia &" and id = " & documento & " and correlativo = "& correlativo
resultado = db.SQLOtras("DBBAECYS",sql)

correlativo = correlativo + 1
sql = "update documento_noticia set correlativo = correlativo - 1 where noticia = "& noticia &" and id <> " & documento & " and correlativo = "& correlativo

resultado = db.SQLOtras("DBBAECYS",sql)

end if

if direccion= "1"  then 'arriba
sql = sql + " set correlativo = correlativo - 1 where noticia = "& noticia &" and id = " & documento & " and correlativo = "& correlativo

resultado = db.SQLOtras("DBBAECYS",sql)
correlativo = correlativo - 1
sql = "update documento_noticia set correlativo = correlativo + 1 where noticia = "& noticia &" and id <> " & documento & " and correlativo = "& correlativo
'Response.Write sql & "<br>"
resultado = db.SQLOtras("DBBAECYS",sql)
'Response.End 

end if


%>
<html>
<script language ="javascript">
function Redireccion(Form) {
redireccion.submit ()
return true
}
</script>
<body onload ="return Redireccion(redireccion)">
<form name="redireccion" action="actualizar_noticia.asp" method="post">
<input name="noticia" value ="<%= noticia%>" type="hidden">
</form>
</body>
</html>