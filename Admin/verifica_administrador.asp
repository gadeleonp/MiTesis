<%@ Language=VBScript %>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>

<%

  set objeto = Server.CreateObject("SQLComandos.Comandos")

  
  
dim login
dim password,password2
dim sql
 login = Request.Form("login")
 password = Request.Form("password")
 repassword = Request.Form("password2")
  
 
 sql = " select id, nombre from administrador "
 Sql = Sql & " where login = '" & login & "' and password = '"& password &"'"
 'Response.Write sql
 'Response.End 
' Response.Redirect ("./Error.asp?error="&sql)
 dim rs 

 
 set rs = objeto.SQLSelect ("DBBAECYS",sql)

 
 if rs.EOF = true then
 %>
 <script language ="javascript">
  function Cambiar(form) {
    form.submit()
    return true
  }
 </script>
 <body onload = "return Cambiar(valida)">
 <form name="valida" action = "Login.asp" method = "post">
  <input id = "error" name = "error" type="hidden" value = "1" >
 </form>
 </body>
 </html>
 <%
 else
 session("administrador") = rs("id")
 Response.Redirect("opciones.asp")
 end if
 set rs = nothing
 'db.close
%>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
