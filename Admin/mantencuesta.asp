<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<HTML>
<HEAD>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<div align="center" >
<TABLE BORDER=0 cellpadding="2" cellspacing="2" >
	<TR>
		<Th class="celdaBgBlue">OPCIONES</Th>
	</TR>

	<TR>
		<TD class="celdaBgBlue05"><A HREF="./crear_pregunta.asp">Crear pregunta</A></TD>
	</TR>
	<TR>
		<TD class="celdaBgBlue05"><A HREF="./ver_encuestas.asp">Ver Encuestas</A></TD>
	</TR>
	<!--
	<TR>
		<TD class="celdaBgBlue05"><A HREF="./publicar_encuesta.asp">Publicar Encuesta</A></TD>
	</TR>
	-->
	<TR>
		<TD class="celdaBgBlue05"><A HREF="./opciones.asp">Regresar al menu anterior</A></TD>
	</TR>
</TABLE>
</div>
</BODY>
</HTML>
