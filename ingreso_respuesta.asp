<%@ Language=VBScript %>
<%
if isempty(session("id")) = true then
  Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim resultado
titulo = Request.Form("titulo")
descripcion = Request.Form("descripcion")
accion = request.form("accion")
foro = Request.Form("foro")
respuesta = request.form("respuesta")
fecha = now
set db = server.CreateObject("SQLComandos.Comandos")
if accion = 1 then
sql = "update respuesta set titulo = '"&titulo & "', descripcion = '"&descripcion&"' where foro = "&foro & " and id = " & respuesta
else
sql = "insert into respuesta (foro,titulo,descripcion,fecha,enlace,autor) values ("& foro & ",'"&titulo &"','"&descripcion &"','" & fecha  & "','',"&session("id")&")"
end if

resultado = db.SQLOtras("DBBAECYS",sql)
set db = nothing

%>
<html>
<script language="javascript">
  function Regresar() {
  pagina = 'desarrollo_foro.asp'
  forma.action = pagina
  forma.submit()
  }
</script>
<body onload="javascript:Regresar()">
<form name="forma" method="get" action="">
<input name="foro" type="hidden" value="<%=foro%>">
</form>
</body>
</html>
