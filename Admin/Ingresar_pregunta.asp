<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>

<%
session("pregunta") = Request.Form("pregunta")
session("respuestas") = Request.Form("respuestas")
session("publicar") = Request.Form("publicar")
%>
<script language="javascript">
</script>
<BODY>
<div align="center" >
<form name="forma" method="post" action="ingresar_encuesta.asp" >
<table border="1" cellspacing="0" cellspanding = "0">
<%
if session("respuestas") > 1 then
   for cont = 1 to session("respuestas") 
   %>
<tr>   
<td class="celdaBgBlueLeft">Respuesta:</td>
<td class="celdaBgBlue05"><input name="respuesta<%=cont%>" value="" type="text" ></td>
</tr>
   <%
   next
   %>
</table>
<input name="tipo_encuesta" type="hidden" value="2">
<input name="Aceptar" type="submit" value="Aceptar">
</form>
</div>
</BODY>
</HTML>
<%
 else
Response.Redirect("ingresar_encuesta.asp?tipo_respuesta=1")
end if
%>
