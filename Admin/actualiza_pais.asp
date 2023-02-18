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
set db = server.CreateObject ("SQLComandos.Comandos")
dim resultado
id = Request.Form("id")
nombre = Request.Form("nombre")
sql = " update pais set nombre = '"+ nombre +"' where id = "+id
resultado = db.SQLOtras ("DBBAECYS",sql)
set db = nothing
%>
<script language="javascript">
 function Continuar(Forma) {
  Forma.submit()
 return true
 }
</script>
<BODY onload="return Continuar(form)">
<FORM NAME="form" ACTION="actualizar_pais.asp" method="post">
<input name="pais" type="hidden" value="<%=id%>">
</FORM>
</BODY>
</HTML>
