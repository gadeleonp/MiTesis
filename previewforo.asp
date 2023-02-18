<%@ Language=VBScript %>
<%
if isempty(session("id")) then
%>
<HTML>
<HEAD>
<TITLE>Previsualizar</TITLE>
</HEAD>
<script language="javascript">
 function Cerrar() {
 opener.location = "Login.asp"
 window.parent.close()
  }
</script>
<BODY onload="javascript:Cerrar()">
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
<link href="theme/Master.css" rel="stylesheet" type="text/css">
</HEAD>
<script language="javascript">
 function Cerrar() {
 window.parent.close()
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
<input name="cerrar" value="Cerrar" type="button" onclick="javascript:Cerrar()">
</div>
</BODY>
</HTML>
<%
end if
%>