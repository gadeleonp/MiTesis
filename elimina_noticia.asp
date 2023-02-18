<%@ Language=VBScript %>
<%
dim db
Dim noticia
Dim Sql
Dim rs
dim resultado
dim objUpload


noticia = Request.Form("noticia")
Sql = " select id, folder_contenedor from noticia where id in (" & noticia &")"

set db = server.CreateObject ("SQLComandos.Comandos")
set objUpload = Server.CreateObject("Dundas.Upload.2")
set rs = db.SQLSelect("DBBAECYS",sql)

while rs.EOF <> true 
    'path = "C:\Documents and Settings\Administrator\My Documents\ProyectoAecys\ProyectoAecys_Local\Noticias\"
    path = server.MapPath("..") & "/Noticias/"
    folder = rs("folder_contenedor")
    path = path & folder
    objUpload.DirectoryDelete path,true
    
    
	rs.MoveNext
wend 
set rs = nothing

%>
<HTML>

<HEAD>
</HEAD>
<BODY>
<%response.write "noticia = " & noticia%>
<h2>Noticia Eliminada Exitosamente</h2>
</BODY>
</HTML>
<%

set objUpload = nothing
Sql = "delete from documento_noticia where noticia in ("& noticia &")"
resultado = db.SQLOtras ("DBBAECYS",sql)
Sql = "delete from noticia where id in ("& noticia &")"
resultado = db.SQLOtras ("DBBAECYS",sql)

Response.Redirect "mantnoticias.asp"
%>