<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
dim db
dim sql
dim rs
dim pass1
dim pass2

  pass1 = request.form("password1")
  pass2 = request.form("password2")

   set db = server.createobject("SqlComandos.Comandos")
   sql  = "select id,login,password from usuario where id = " & session("id")  &"  and password = '" & pass1 & "'"

	response.write sql
   set rs = db.SQLSelect("DBBAECYS",sql)
	
	if not rs.eof then
	   password = rs("password")

	   	sql = "update usuario set recibir_mail =  '" & NO & "', habilitada = 0 where id = " & session("id")
		resultado = db.sqlotras("DBBAECYS",sql)
		
	else
     response.redirect("eliminar_cuenta.asp?Error=1")
	end if
	

	
	
	set rs = nothing
	set db = nothing
	

%>