<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<%
dim db  
dim login
dim password,password2
dim sql
dim empresa
dim rs 
dim resultado
 
 login = Request.Form("txtlogin")
 password = Request.Form("txtpassword")

 set db = server.CreateObject ("SQLComandos.Comandos")
 
  
 
 sql = " select id, nombre,apellido,login,password,empresa from usuario "
 Sql = Sql & " where login = '" & login & "' and password = '"& password &"' and habilitada = 1"
 
 set rs = db.SQLSelect ("DBBAECYS",sql)
	if rs.EOF = true then
		Response.redirect("Login.asp?Error=1")
	else
	empresa = ""
	empresa = rs("empresa")
	
	if (instr(1,rs("password"),password) > 0) then
	
		if isnull(empresa)then
			session("id") = rs("id")
			session("nombre") = rs("nombre")+" "+rs("apellido")
			sql = "update usuario set ultima_visita = '"& now & "' where id = " & rs("id")
			
			resultado = db.SQLOtras ("DBBAECYS",sql)
			if (resultado = -1) then
			  Response.Write "Error al actualizar los datos"
			end if
		else
			session("UsEmpresa") = rs("id")
			session("UsEmpNombre") = rs("nombre")+" "+rs("apellido")
		end if
   else
	Response.redirect("Login.asp?Error=1")
   end if

 if len(rs("empresa")) > 0 then
 empresa = 1
 else
 empresa = 0
 end if 
 
 end if
 set rs = nothing
 
 
 if empresa = 0 then
%>
<script language="javascript">
function Redireccion() {
destino = "./principal.asp?ver=0&path=opciones_datos.asp"
window.parent.location = destino
}
</script>
<body onload="return Redireccion()">
</body>
</HTML>
<%
else
%>
<script language="javascript">
function Redireccion() {
destino = "./principal.asp?ver=0&path=./empresa/opciones_empresa.asp"
window.parent.location = destino
}
</script>
<body onload="return Redireccion()">
</body>
</HTML>
<%
end if
%>