<%@ Language=VBScript %>
<%
if isempty(session("UsEmpresa")) then
%>
<HTML>
<HEAD>
<TITLE>Previsualizar</TITLE>
</HEAD>
<script language="javascript">
 function Aceptar() {
 window.location = "../Login.asp"
 
  }
</script>
<BODY onload="javascript:Aceptar()">
</BODY>
</HTML>
<%
else
%>
<%
dim Desc
desc = Request.QueryString("desc")
%>
<HTML>
<HEAD>
<TITLE>Previsualizar</TITLE>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
 function Aceptar() {
  window.location = "./opciones_empresa.asp"
  }
</script>
<BODY>
<table border="0">
<tr>
<td>
<%
function FormatStr(fString)
	on Error resume next
    fString = Replace(fString, CHR(13) & CHR(10), "<br>")
	FormatStr = fString
	on Error goto 0
end function
%>
<%
Response.Write desc
%>
</td>
</tr>
<table>
<div align="center">
<input name="cerrar" value="Aceptar" type="button" onclick="javascript:Aceptar()">
</div>
</BODY>
</HTML>
<%
end if
%>