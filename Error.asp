<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
hola
<%
dim error
error = Request.Form("error")
Response.Write("error = "+error)
 %>
</BODY>
</HTML>
