<%@ Language=VBScript %>
<%
if isempty (session("administrador")) then
 Response.Redirect("Login.asp")
end if
%>
<html>
<head>
<link href="../theme/Master.css" rel="stylesheet" type="text/css">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script language="javascript">
 function Regresar() {
 window.location="mantencuesta.asp"
 }
 function Grafica(pregunta) {
	  ancho = 320
	  altura = 450
	  if (screen) {
	  leftPos = (screen.width / 2) - (ancho/2)
	  }
	  NuevaVentana = window.open('../grafico.asp?pregunta='+pregunta,'Nueva_Encuesta','toolbar=no,menubar=no,location=no,scrollbars=no,width='+ancho+'+,height='+altura+',left='+leftPos+'+,top=20')
 }
</script>

<%
dim db
dim rs
dim rs2
dim sql
set db = server.CreateObject("SQLComandos.Comandos")
sql = "select id,descripcion,day(fecha) dia, month(fecha) mes, year(fecha) anio from pregunta order by fecha desc"
set rs = db.SQLSelect("DBBAECYS",sql)
%>
<body>
<%
if not rs.eof then
%>
<h1 align="center">Reporte de Encuestas</h1><br>

<table width="0%" border="0" align="center" cellpadding="2" cellspacing="2">
<tr align="center">
    <th height="25" class="celdaBgBlue">Preguntas</th>
  </tr>
<%

while rs.eof <> true
sql = "select opr.descripcion, opr.total from opciones_respuesta opr where opr.pregunta =" & rs("id")
set rs2 = db.SQLSelect("DBBAECYS",sql)

%>

  <tr>
    <td class="celdaBgBlue05"><%=rs("descripcion")%></td>
    <td><%=rs("dia")%>/<%=rs("mes")%>/<%=rs("anio")%></td>	
	<%
	
	if not rs2.eof then
	cont_opcion = 1
	%>
<tr>
<td>
	<table border="0" cellpadding="2" cellspacing="2">
        <tr class="celdaBgBlue"><th><div align="center">No. </div></th>
        <th><div align="center">Opcion </div></th>
        <th><div align="center">Total </div></th>
        <%while  rs2.eof <> true%>
        <tr class="celdaBgBlue05"> 
          <td><%=cont_opcion%> <div align="center"></div></td>
          <td><%=rs2("descripcion")%> <div align="center"></div></td>
          <td><%=rs2("total")%> <div align="center"></div></td>
        </tr>
        <%
	cont_opcion = cont_opcion + 1
	rs2.movenext
	%>
        <%wend%>
      </table>	
</td>
<td>
<img alt="Ver grafico" onClick="javascript:Grafica('<%=rs("id")%>')" src="iconos/refresh.gif" WIDTH="20" HEIGHT="20"></a>
</td>
</tr>	
	<%	
	else
	%>	
	<h3>No han habido votaciones</h3>
	<%	
	end if
	set rs2 = nothing
	%>	

  </tr>
<%
rs.movenext
wend%>  
</table>

<%
else
%>
<h2 align="center">No se han creado encuestas</h2>

<%
end if 
set rs = nothing
set db = nothing
%>
<br>
<div align="center">
  <input name="Regresar" type="button" class="celdaBgBlue05" onclick="javascript:Regresar()" value="Regresar">
</div>
</body>
</html>
