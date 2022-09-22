<%@ Language=VBScript %>

<html>
<head>
<%
dim menu
menu = ""
menu = request.QueryString("menu")

if not isempty(session("id")) then
	if menu = "1" then
    Response.Redirect("opciones_datos.asp")
	end if
	if menu="2" then
    Response.Redirect("foro.asp")
	end if
end if

dim Error
Error = ""
Error = Request.Form ("Error")


%>
<style type="text/css">
<!--
boton {
	background-color: #ffffff;
	border-top: thin none #ffffff;
	border-right: thin none;
	border-bottom: thin none;
	border-left: thin none;
	font-family: Arial, Helvetica, sans-serif;
	font-style: normal;
	line-height: normal;
	font-weight: bold;
	font-variant: normal;
	text-decoration: underline;
}
-->
</style>
</head>
<script language="javascript">
function Ingresar() {
if (verificar(document.FORM1) == true) {
FORM1.submit()
}
}
 function verificar(form) {

   if (form.txtlogin.value.length == 0) {
   alert("No ingreso su login")
   form.txtlogin.select()
   form.txtlogin.focus()
   return false
   }
   if (form.txtpassword.value.length ==0) {
   alert("No ingreso su contraseña")
   form.txtlogin.select()
   form.txtlogin.focus()
   return false
   }
 
  return true
 }
</script>
<body>
<div align="center">
<h2>
<%
Response.Write(Error)
if error = "1" then
Response.Write(error)
end if
%>
</h2>

<form onsubmit="return verificar(this)" id="FORM1" name="FORM" action="./verifica_usuario.asp" method="post">
    <table width="180" height="54" border="1" bordercolor="#333333" bgcolor="#0099FF">
      <tr> 
        <td width="70"><div align="right"><strong><font color="#333333" size="2" face="Arial, Helvetica, sans-serif">Login:</font></strong></div></td>
        <td width="116"><div align="center"> 
            <input name="txtlogin" style="FONT:Georgia WIDTH: 60px height:40px" type="text" size="15" maxlength="15">
          </div></td>
      </tr>
      <tr> 
        <td><div align="right"><strong><font color="#333333" size="2" face="Arial, Helvetica, sans-serif">Password:</font></strong></div></td>
        <td><div align="center"> 
            <input name="txtpassword" style="FONT:Georgia WIDTH: 60px height:40px" type="password" size="15" maxlength="15">
          </div></td>
      </tr>
    </table>
	<br>
	<input  class="boton" name="entrar" value="Login" type="submit">
  </form>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
  <font color="#003399" size="3" face="Arial, Helvetica, sans-serif">

  </font> 
</div>
</body>
</html>
