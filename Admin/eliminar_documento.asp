<%@ Language=VBScript %>
<%
dim db
dim noticia
dim documento
dim objupload
dim resultado

noticia = Request.form("noticia")
documento = Request.form("documento") 
	if len(noticia) = 0 then
		noticia = request.Querystring("noticia")
		documento = request.QueryString("documento")
	end if
set objupload = server.CreateObject ("dundas.upload.2")
set db = server.CreateObject("SQLComandos.Comandos")

sql = "select folder_contenedor from noticia where id = " & noticia

dim rs
set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.eof then
	folder = rs("folder_contenedor")
end if

set rs = nothing
sql = "select id,enlace, correlativo from Documento_Noticia where noticia = " & noticia & " and id = " & documento

set rs = db.SQLSelect("DBBAECYS",sql)

if not rs.eof then
	id_documento = rs("id")
	documento = rs("enlace")
	correlativo = rs("correlativo")
end if
set rs = nothing

objupload.UseVirtualDir = true
path =  "/aecys/paecycod/Noticias/" & folder & "/"

sql = "select count(correlativo) NumDocs from documento_noticia where noticia = " & noticia
set rs = db.SQLSelect("DBBAECYS",sql)
if not rs.eof then
NumDoc = rs("NumDocs")
end if
set rs = nothing

'Response.Write "documento = " & documento & "<br>"
'Response.Write "Correlativo  = " & correlativo & "<br>"
'Response.Write "NumDocs = " & NumDoc & "<br>"

if correlativo = 1 and NumDoc > 1 then
   sql = "delete from documento_noticia where id = " & id_documento 
	resultado = db.SQLOtras("DBBAECYS",sql)
   sql = "update documento_noticia set correlativo = correlativo - 1 where noticia = " & noticia & " and correlativo > " & correlativo  
	resultado = db.SQLOtras("DBBAECYS",sql)
	objupload.FileDelete (path+documento)
end if

if correlativo = NumDoc and NumDoc > 1 then
	sql = "delete from documento_noticia where id = " & id_documento
	resultado = db.SQLOtras("DBBAECYS",sql)
	objupload.FileDelete (path+documento)
end if

if correlativo > 1 and correlativo < NumDoc then
    sql = "delete from documento_noticia where id = " & id_documento
	resultado = db.SQLOtras("DBBAECYS",sql)
	objupload.FileDelete (path+documento)
    sql = "update documento_noticia set correlativo = correlativo - 1 where noticia = " & noticia & " and correlativo > " & correlativo 
	resultado = db.SQLOtras("DBBAECYS",sql)
end if

'Response.Write sql
'Response.Write path+documento

set objupload = nothing
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script language = "javascript" >
function Regresar(Form) {
Form.submit()
return true
}
</script>
<BODY onload = "return Regresar(regresar)">
<form name="regresar" action = "actualizar_noticia.asp" method="post">
<input name="noticia" type = "hidden" value ="<%=noticia%>">
</form>


</BODY>
</HTML>
