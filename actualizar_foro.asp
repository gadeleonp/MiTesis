<%@ Language=VBScript %>
<%
if isempty(session("id")) then
  Response.Redirect("login.asp")
end if
dim db
dim resultado
Dim sql
dim titulo
dim descripcion
dim foro 
set db = server.CreateObject("SQLComandos.Comandos")
foro = Request.Form("foro")
titulo = Request.Form("titulo")
categoria = Request.Form("categoria")
descripcion = Request.Form("descripcion")
sql = "update foro set descripcion = '" & descripcion &"', titulo = '"& titulo &"', categoria = "&categoria&" where id = "& foro


resultado = db.SQLOtras("DBBAECYS",sql)
set db = nothing

Response.redirect(session("destforo")+"?foro="+foro)

%>