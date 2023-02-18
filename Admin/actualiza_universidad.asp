<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<H2>DATOS ACTUALIZADOS CORRECTAMENTE</H2>
</HEAD>
<%
dim db
dim rs
dim resultado
set db = server.CreateObject ("SQLComandos.Comandos")
id = Request.Form("id")
nombre = Request.Form("nombre")
direccion = Request.Form("direccion")
telefono = Request.Form("telefono")
sitio = Request.Form("sitio")
email = Request.Form("email")
pais = Request.Form("pais")

sql = " update universidad set nombre = '"+ nombre +"' ,direccion = '"+direccion +"' ,telefono = '"+telefono+"', sitio = '"+sitio+"' , email = '"+email+"',pais = "+ pais + " where id = "+id
resultado = db.SQLOtras("DBBAECYS",  sql)

%>
<script language="javascript">
 function Continuar(Forma) {
  Forma.submit()
 return true
 }
</script>
<BODY onload="return Continuar(form)">
<FORM NAME="form" ACTION="actualizar_universidad.asp" method="post">
<input name="universidad" type="hidden" value="<%=id%>">
</FORM>
</BODY>
</HTML>
