<%@ Language=VBScript %>
<HTML>
<HEAD>
<%
dim Error
Error = ""
Error = Request.Form ("Error")
%>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">

<script language="javascript">
 function verificar(form) {

   if (form.login.value.length == 0) {
   alert("No ingreso su login")
   return false
   }
   if (form.password.value.length ==0) {
   alert("No ingreso su contraseña")
   return false
   }
 
  return true
 }
</script>
</HEAD>
<BODY>
<div align="center">
<h2>
<%

if error = "1" then
Response.Write("LOGIN INCORRECTO")
error = ""
end if
%>
</h2>
<FORM onsubmit="return verificar(this)" id="FORM1" name="FORM" action="verifica_administrador.asp" method="post">
    <TABLE WIDTH="24%" BORDER="0" cellpadding="2" cellspacing="2" >
      <TR>
		<th width="33%" align="right" nowrap class="celdaBgBlueLeft" ><div align="center">Login:</div></th>
		<TD width="67%"> 
          <INPUT id="login" name="login" type="text">
		</TD>
	</TR>
	<TR>
		<th width="33%" align="right" nowrap class="celdaBgBlueLeft" ><div align="center">Password: </div></th>
		<TD width="67%"> 
          <INPUT id=¨"password" type="password" name="password"></TD>
	</TR>
	</TABLE>
	<INPUT id="Aceptar" type="submit" name="aceptar" value ="Aceptar">
</FORM>
</div>
</BODY>
</HTML>
