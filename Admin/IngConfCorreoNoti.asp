<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<%
dim db
dim sql

set db = server.CreateObject("SqlComandos.Comandos")
 
 'Parametros de configuracion de correo
 tipocorreo = request.form("tipocorreo")
 publicar = Request.Form("chkenviacorreo")
 autor = Request.Form("autor")
 titulo = Request.Form("titulo")
 resumen = Request.Form("resumen")
 tema = Request.Form("tema")
 saludo = Request.Form("saludo")
 tipoletra = Request.Form("tipoletra")
 negrilla = Request.Form("negrilla")
 correo = Request.Form("correo")

 tituloautorresumen = titulo & autor & resumen 'Informacion a concatener al correo

 'Actualiza la tabla de configuracion de correos
	if tipocorreo = "Noticia" then
 		sql = "update ConfCorreo set publicar = '" & publicar & "', tituloautorresumen = '" &tituloautorresumen& "', subject = '"&tema &"', saludo = '" &saludo & "', tipoletra = "&tipoletra & ",negrilla = '" & negrilla & "', cuenta = '"&correo&"' where descripcion = 'Noticia'"
		destino = "ConfCorreoNoti.asp"
 	end if
	 if tipocorreo = "Foro" then
		sql = "update ConfCorreo set publicar = '" & publicar & "', tituloautorresumen = '" &tituloautorresumen& "', subject = '"&tema &"', saludo = '" &saludo & "', tipoletra = "&tipoletra & ",negrilla = '" & negrilla & "', cuenta = '"&correo&"' where descripcion = 'Foro'"	
		destino = "ConfCorreoForo.asp"
	end if
	 if tipocorreo = "Empresa" then
		sql = "update ConfCorreo set publicar = '" & publicar & "', tituloautorresumen = '" &tituloautorresumen& "', subject = '"&tema &"', saludo = '" &saludo & "', negrilla = '" & negrilla & "', cuenta = '"&correo&"' where descripcion = 'Empresa'"	
		destino = "ConfCorreoEmp.asp"
	end if
 resultado = db.SQLOtras ("DBBAECYS",sql)
 set db = nothing
 Response.Redirect (destino)
%>
