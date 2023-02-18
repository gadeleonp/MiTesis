<%@ Language=VBScript %>
<%
if isempty(session("id")) then
    Response.Redirect("Login.asp")
end if

Dim id
dim login
id = session("id")
'login = Request.QueryString ("login")


%>
<HTML>
<HEAD>
<title>
Ingreso de Curriculum
</title>
<link href="./theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
<%
dim menu
menu = request.QueryString("menu")
if menu <> "1" then
%>
function cerrarventana() {
	   window.close()
 	}
<%
end if
%>	
function Validar(Forma) {
if ((Forma.curriculum.value.length == 0) && (Forma.action == "ingresar_curriculum.asp"))
{
	alert ("No ingreso archivo")
	return false
}
return true
}


</script>
<BODY onload="javascript:{if(parent.frames[0]&&parent.frames['opciones'].Go)parent.frames['opciones'].Go()}" >
<%if menu <> "1" then%>

<%end if%>
<div align="center">
<form  name="ingcurriculum" method="post" enctype="multipart/form-data" action="ingresar_curriculum.asp" >
<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td class="celdaBgBlueLeft">Curriculum: </td>
<td class="celdaBgBlue05">
<input name="curriculum" type="file">
<input name="menu" type="hidden" value="<%=menu%>">
</td>
</tr>
</table>
<br>

<table border="0" cellpadding="2" cellspacing="2">
<tr>
<td><input name="aceptar" type="submit" class="celdaBgBlue05" onClick="return Validar(ingcurriculum)" value="Aceptar"> </td>
<%if menu <> "1" then%>
<td width="17"></td>
<td><input name="cancelar" type="submit" class="celdaBgBlue05" onClick="return cerrarventana()" value="Cancelar"></td>
<%end if%>
</tr>
</table>


</form>
</div>
</BODY>
</HTML>
