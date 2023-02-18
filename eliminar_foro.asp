<%@ Language=VBScript %>
<%
if isempty(session("id")) then
	Response.Redirect("login.asp")
end if
dim db
dim resultado
dim foro
foro = Request.QueryString ("foro")
accion = request.form("accion")
set db = server.CreateObject("SQLComandos.Comandos")

if not isempty(accion) then
	foros = request.form("publicar")
	sql = "update foro set activo = 1 where  autor = " &session("id") &" and id in ("& foros &")"
	resultado = db.SQLOtras("DBBAECYS",sql)
	response.redirect(session("destforo"))

	else
	
'******************* Eliminar los documentos que hayan sido ingresados al foro	
dim ObjUpload
dim rs
set ObjUpload = server.CreateObject("dundas.upload.2")

	sql = "select archivo from documento where foro ="&foro
 
  
  set rs = db.SQLSelect("DBBAECYS",sql)
	if not rs.EOF then
		while rs.eof <> true
			archivo =  rs("archivo")
			RootFolderName = Server.MapPath(".") & "\"
			Path = RootFolderName & "Foros" &"\" & archivo
			objupload.FileDelete (Path)
			rs.movenext
	    wend
	end if
	set rs = nothing

'*******************
	sql = "delete from documento where foro = "&foro
	resultado = db.SQLOtras("DBBAECYS",sql)

	
	sql = "delete from foro where id = "& foro &" and autor = "& session("id")
	resultado = db.SQLOtras("DBBAECYS",sql)


'*******************	

response.redirect(session("destforo"))
end if
   set db = nothing
%>