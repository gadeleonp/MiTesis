<%@ Language=VBScript %>
<%
if isempty(session("id")) = true then
  Response.Redirect("Login.asp")
end if
%>
<%
dim db 
set db = server.CreateObject("SQLComandos.Comandos")
dim resultado
titulo = Request.Form("titulo")
descripcion = Request.Form("descripcion")
categoria = Request.Form("categoria")
fecha = now

sql = "insert into foro (autor,titulo,descripcion,activo,fecha_creacion,estado,categoria) values ("& session("id") & ",'"&titulo &"','"&descripcion &"',0,'"& fecha &"',1,"&categoria&")"
resultado = db.SQLOtras("DBBAECYS",sql)
set db = nothing

Response.redirect("reporte_foros.asp")
%>
